# -*- coding: utf-8 -*-
"""BỘ TOOL CHỈ ĐỌC cho bot TÌM LỖI (`tim_loi`) và NGHIỆM THU (`nghiem_thu`) — 15/10/2026, hợp đồng đợt 2+3 mục C
(`wt/hopdong_dot2_agent_tong_20261015.md`).

Bot tìm lỗi trước đây chạy bằng Claude Code (Bash tự do: openpyxl, psql, gọi builder). Trong chat (OpenClaw) bot không
có exec, nên mọi bước đó thành tool TẤT ĐỊNH ở đây. Hai nhóm:

  · 8 tool LẺ (hợp đồng C): tl_doc_mapping · tl_tim_file_nguon · tl_doc_excel · tl_tong_hop_excel · tl_raw_rows ·
    tl_so_dashboard · tl_doc_code — phần Excel ở `tim_loi_excel.py`.
  · 4 tool GỘP BƯỚC (bổ sung của giám sát 15/10): tl_ho_so_man · tl_kiem_nhanh_nap · tl_so_qua_ky · tl_truy_vet —
    mỗi tool thay 5-8 lệnh bot tìm lỗi vẫn gõ tay. SKILL bảo bot gọi chúng TRƯỚC, chỉ xuống tool lẻ khi cần.

AN TOÀN (bất biến, có test ở scripts/test_timloi_eval.py):
  · DB: kết nối của `so_lieu_check._ket_noi` — `default_transaction_read_only=on` + statement_timeout; env chỉ
    prod (5434) | test (5435). KHÔNG có tool nhận SQL: câu SQL viết sẵn, tham số đi qua placeholder; report_type
    phải khớp `_RE_RT`, ngày phải ISO thật, cột lọc/nhóm chỉ trong `_COT_DB` hoặc `payload.<khoá>` (lọc payload làm
    ở Python, không ghép vào SQL).
  · File: chỉ trong received_reports / Tài liệu, chặn `..`/symlink (xem `tim_loi_excel.duong_an_toan`); mã nguồn chỉ
    đọc .py dưới `app/` của tc-admin-api.
  · Builder: chạy trong tiến trình con bằng venv tc-admin-api (`scripts/timloi_so_dump.py`), DB chỉ đọc qua PGOPTIONS.

CACHE (dữ liệu có thể CŨ TỚI 5 PHÚT — mọi tool có cache nhận `bo_cache=true` để đo lại):
  · kết quả builder theo (env, màn, kỳ, grain, khối, công ty, từ/đến ngày) — 5 phút;
  · workbook theo (đường dẫn, mtime) — LRU (tim_loi_excel);
  · hồ sơ màn theo mtime spec/builder/mapping/FE — tự hết hạn khi một trong các file đó đổi.
"""
import datetime
import fnmatch
import glob
import json
import os
import re
import subprocess
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal

from . import man_fe as mf
from . import mapping_check as mc
from . import so_lieu_check as sc
from . import tim_loi_excel as tx

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
SPEC_DIR = os.path.join(_AGENT_ROOT, "extract_specs")
DUMP = os.path.join(_AGENT_ROOT, "scripts", "timloi_so_dump.py")
ROOT = os.environ.get("TC_ROOT") or "/home/itadmin/AI_Dashboard_QT"
LOGS_DIR = os.environ.get("TC_LOGS_DIR") or os.path.join(ROOT, "AI_coding", "logs")
META = os.environ.get("AVAILABLE_METADATA") or os.path.join(ROOT, "Connect_VPS", "available_metadata.json")
API_DIR = {"prod": mc.API_DIR, "test": os.environ.get("TC_API_DIR_TEST") or os.path.join(ROOT, "AI_coding", "tc-admin-api")}
_PORT = {"prod": "5434", "test": "5435"}
_VN = datetime.timezone(datetime.timedelta(hours=7))

TTL_BUILDER = 300
_RE_RT = re.compile(r"^[A-Z][A-Z0-9_]{1,47}$")
_RE_MAN = re.compile(r"^[a-z][a-z0-9_]{1,23}$")
_RE_HAM = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,80}$")
_RE_PAYLOAD = re.compile(r"^payload\.([A-Za-z_][A-Za-z0-9_]{0,40})$")
_COT_DB = ("source_file", "ngay", "period_month", "cong_ty", "khoi", "cost_center", "dim1", "dim2", "dim3", "dataset_kind")
MAX_DONG_DB = 50000
_RE_LOI_CRON = re.compile(r"rc=(?!0\b)-?\d+|traceback|exception|\blỗi\b|\bloi\b|\bdừng\b|error", re.IGNORECASE)


def _env(env):
    if env not in _PORT:
        raise ValueError("env phải là 'prod' hoặc 'test'")
    return env


def _ngay(s, ten):
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(s or "")):
        raise ValueError(f"{ten} phải dạng YYYY-MM-DD")
    datetime.date.fromisoformat(s)                      # ngày không tồn tại (2026-02-30) -> ValueError
    return s


def _ky_sang_khoang(ky):
    """'YYYY-MM' -> (ngày đầu, ngày cuối); 'YYYY-MM-DD' -> (ngày, ngày)."""
    if re.fullmatch(r"\d{4}-\d{2}", str(ky or "")):
        y, m = map(int, ky.split("-"))
        if not 1 <= m <= 12:
            raise ValueError("tháng không hợp lệ")
        cuoi = datetime.date(y + m // 12, m % 12 + 1, 1) - datetime.timedelta(days=1)
        return f"{ky}-01", cuoi.isoformat()
    d = _ngay(ky, "ky")
    return d, d


def _dec(x):
    return Decimal(repr(float(x))) if x is not None else Decimal(0)


def _so(d):
    return int(d) if d == d.to_integral_value() else float(d)


def _gio_vn(ts):
    return datetime.datetime.fromtimestamp(ts, tz=_VN).strftime("%Y-%m-%d %H:%M:%S")


def _rt_hop_le(rt):
    if not _RE_RT.fullmatch(str(rt or "")):
        raise ValueError("report_type chỉ gồm chữ HOA/số/_ (vd HTXT_GIAO)")
    return rt


def _an(env):
    return {r["source_file"] for r in sc._ro(env, "SELECT source_file FROM hidden_files")}


def _phu_luc_bang_chung(ten_tool, tham_so):
    """Chuỗi 'lệnh' để bot dán vào bang_chung[].lenh — người khác chạy lại được đúng phép đo."""
    return f"{ten_tool}({json.dumps(tham_so, ensure_ascii=False, default=str)})"


# ═════════════════════════════ 1. MAPPING ═════════════════════════════
def _ban_copy(file):
    """'5.HTxetai_Mapping_Dashboard copy.xlsx' <-> '5.HTxetai_Mapping_Dashboard.xlsx' — bẫy BẢN COPY mapping: hai bản
    cùng tồn tại, chỉ một bản được cập nhật. Trả các bản 'anh em' kèm mtime để bot nói rõ đang đọc bản nào."""
    goc = re.sub(r"\s*copy(\s*\(\d+\))?(?=\.xls)", "", file, flags=re.IGNORECASE)
    stem = os.path.splitext(goc)[0]
    ra = []
    for p in glob.glob(os.path.join(glob.escape(tx.TAI_LIEU_DIR), "*.xls*")):
        b = os.path.basename(p)
        if os.path.splitext(re.sub(r"\s*copy(\s*\(\d+\))?(?=\.xls)", "", b, flags=re.IGNORECASE))[0] == stem:
            ra.append({"file": b, "sua_luc": _gio_vn(os.path.getmtime(p)), "dang_doc": b == file})
    return sorted(ra, key=lambda x: x["sua_luc"], reverse=True)


def _dong_tieu_de_mapping(rows):
    for i, r in enumerate(rows[:12]):
        if any(re.match(r"^chi tieu|truong tren giao dien", tx._kd(c)) for c in r if c is not None):
            return i + 1
    return None


def doc_mapping(tai_lieu=None, file=None, sheet=None, dong=None, den_dong=None, tu_khoa=None):
    """Dòng mapping NGUYÊN VĂN, mọi cột (cột Excel + tiêu đề + giá trị). Chọn tài liệu bằng mã `tai_lieu` của
    mapping_index (QTTC, VHKD, XDV, DUAN, HTXT, ANTAXI, XANHVP, QLTS) hoặc `file` trong Tài liệu/. Không `sheet` ->
    sheet đầu tiên mapping_index khai (bản mới nhất đã khai) + danh sách sheet. `dong` (+ `den_dong`, tối đa 60 dòng)
    = số dòng Excel như trong log. `tu_khoa` -> các dòng chứa từ khoá (không dấu) để tìm số dòng."""
    ci = mc._chi_muc()
    if tai_lieu:
        if tai_lieu not in ci["tai_lieu"]:
            raise ValueError(f"tài liệu '{tai_lieu}' không có. Có: {', '.join(ci['tai_lieu'])}")
        d = ci["tai_lieu"][tai_lieu]
        file, sheet_khai = d["file"], d["sheet"]
    elif file:
        sheet_khai = next((d["sheet"] for d in ci["tai_lieu"].values() if d["file"] == os.path.basename(file)), [])
    else:
        return {"tai_lieu": {k: {"file": d["file"], "sheet": d["sheet"], "man": d["man"]} for k, d in ci["tai_lieu"].items()}}
    p = tx.duong_an_toan(os.path.join(tx.TAI_LIEU_DIR, os.path.basename(file)) if not os.path.isabs(file) else file)
    tat_ca = tx.ds_sheet(p)
    sheet = sheet or (sheet_khai[0] if sheet_khai else tat_ca[0])
    _, ten, rows = tx.doc_sheet(p, sheet)
    hd = _dong_tieu_de_mapping(rows)
    tieu_de = {j: str(c).strip() for j, c in enumerate(rows[hd - 1]) if c not in (None, "")} if hd else {}

    def dong_ra(i):
        r = rows[i - 1] if i <= len(rows) else []
        return {"dong": i, "cot": [{"cot": tx.chu_cot(j), "tieu_de": tieu_de.get(j), "gia_tri": tx.json_gia_tri(c)}
                                   for j, c in enumerate(r) if c not in (None, "")]}

    def nhom_tren(i):
        for k in range(i - 1, (hd or 0), -1):
            r = [c for c in rows[k - 1] if c not in (None, "")]
            if len(r) == 1 and rows[k - 1] and rows[k - 1][0] not in (None, "") and not str(r[0]).strip().isdigit():
                return {"dong": k, "nhom": str(r[0]).strip()[:200]}
        return None
    out = {"file": os.path.basename(p), "sheet": ten, "sheet_khai_trong_index": sheet_khai, "sheet_co_trong_file": tat_ca,
           "dong_tieu_de": hd, "ban_cung_ten": _ban_copy(os.path.basename(p)),
           "luu_y": "Đối chiếu với bản mapping MỚI NHẤT (xem ban_cung_ten + sheet có ngày mới nhất). 'Dòng N' trong log = "
                    "số dòng Excel của sheet đó."}
    if dong is not None:
        a = int(dong)
        b = min(int(den_dong) if den_dong else a, a + 59)
        if a < 1:
            raise ValueError("dong bắt đầu từ 1")
        out["dong"] = [dong_ra(i) for i in range(a, b + 1)]
        out["nhom_cua_dong_dau"] = nhom_tren(a)
    if tu_khoa:
        kw = tx._kd(tu_khoa)
        if len(kw) < 2:
            raise ValueError("từ khoá quá ngắn")
        khop = [i for i in range((hd or 0) + 1, len(rows) + 1)
                if kw in tx._kd(" ".join(str(c) for c in rows[i - 1] if c is not None))]
        out["dong_chua_tu_khoa"] = [dong_ra(i) for i in khop[:30]]
        out["so_dong_khop"] = len(khop)
    return out


# ═════════════════════════════ 2. FILE NGUỒN ═════════════════════════════
_meta_cache = {}


def _metadata():
    """available_metadata.json (danh mục máy gửi đã quét) -> {(CÔNG TY, tên file): mục}. Cache theo mtime."""
    try:
        mt = os.path.getmtime(META)
    except OSError:
        return {}
    if _meta_cache.get("mt") != mt:
        try:
            with open(META, encoding="utf-8") as fh:
                ds = json.load(fh)
        except (OSError, ValueError):
            ds = []
        idx = {}
        for x in ds if isinstance(ds, list) else []:
            if isinstance(x, dict) and x.get("fileName"):
                idx.setdefault((str(x.get("company") or "").upper(), x["fileName"]), x)
        _meta_cache.update({"mt": mt, "idx": idx})
    return _meta_cache["idx"]


def _sidecar(p):
    j = os.path.splitext(p)[0] + ".json"
    try:
        with open(j, encoding="utf-8") as fh:
            x = json.load(fh)
        return x if isinstance(x, dict) else {}
    except (OSError, ValueError):
        return {}


def _mo_ta_file(p, don_vi):
    st = os.stat(p)
    sc_ = _sidecar(p)
    me = _metadata().get((don_vi.upper(), os.path.basename(p)))
    return {"file": os.path.relpath(p, tx.RECEIVED_DIR), "ten": os.path.basename(p), "co_byte": st.st_size,
            "sua_tren_252": _gio_vn(st.st_mtime), "luu_ve_252": sc_.get("saved_at"),
            "sua_o_may_nguon": sc_.get("modified_at"), "nguoi_sua": sc_.get("last_modified_by"),
            "duong_may_nguon": sc_.get("original_source_path"),
            "trong_available_metadata": bool(me), "metadata_modifiedAt": (me or {}).get("modifiedAt")}


def _thu_muc_con(goc, ten, loai):
    if ten in (None, ""):
        return None
    if "/" in str(ten) or "\\" in str(ten) or str(ten).strip() in (".", ".."):
        raise ValueError(f"{loai} là TÊN một thư mục, không phải đường dẫn")
    for n in sorted(os.listdir(goc)):
        if n.upper() == str(ten).strip().upper() and os.path.isdir(os.path.join(goc, n)) \
                and not os.path.islink(os.path.join(goc, n)):
            return n
    raise ValueError(f"không có {loai} '{ten}'. Có: {', '.join(sorted(x for x in os.listdir(goc) if os.path.isdir(os.path.join(goc, x))))}")


def tim_file_nguon(don_vi, thu_muc=None, tu_khoa=None, gioi_han=60):
    """File nguồn đã về 252 (received_reports/<ĐƠN VỊ>/<thư mục>/), mới nhất trước: mtime (giờ VN), lúc lưu về 252 và
    người sửa ở máy nguồn (file .json đi kèm), có trong available_metadata không. Không `thu_muc` -> liệt kê thư mục con
    kèm số file + file mới nhất. `tu_khoa` lọc theo tên file (không dấu, không phân biệt hoa thường)."""
    gioi_han = max(1, min(int(gioi_han or 60), 300))
    dv = _thu_muc_con(tx.RECEIVED_DIR, don_vi, "đơn vị")
    goc = os.path.join(tx.RECEIVED_DIR, dv)
    kw = tx._kd(tu_khoa) if tu_khoa else None
    if not thu_muc and not kw:
        ra = []
        for n in sorted(os.listdir(goc)):
            d = os.path.join(goc, n)
            if not os.path.isdir(d) or os.path.islink(d):
                continue
            fs = [p for p in glob.glob(os.path.join(glob.escape(d), "*.xls*")) if os.path.isfile(p)]
            moi = max(fs, key=os.path.getmtime) if fs else None
            ra.append({"thu_muc": n, "so_file": len(fs),
                       "file_moi_nhat": os.path.basename(moi) if moi else None,
                       "sua_luc": _gio_vn(os.path.getmtime(moi)) if moi else None})
        return {"don_vi": dv, "thu_muc": ra}
    tms = [_thu_muc_con(goc, thu_muc, "thư mục")] if thu_muc else \
        [n for n in sorted(os.listdir(goc)) if os.path.isdir(os.path.join(goc, n)) and not os.path.islink(os.path.join(goc, n))]
    fs = []
    for tm in tms:
        for p in glob.glob(os.path.join(glob.escape(os.path.join(goc, tm)), "*.xls*")):
            if os.path.isfile(p) and not os.path.islink(p) and (not kw or kw in tx._kd(os.path.basename(p))):
                fs.append(p)
    fs.sort(key=os.path.getmtime, reverse=True)
    return {"don_vi": dv, "thu_muc": thu_muc and tms[0], "so_file": len(fs),
            "file": [_mo_ta_file(p, dv) for p in fs[:gioi_han]], "cat_bot": max(0, len(fs) - gioi_han),
            "luu_y": "Tên file nhiều khi KHÔNG lăn tháng (file luỹ kế ghi đè) và có thể có 2 tên cùng nội dung — "
                     "xem sua_o_may_nguon/luu_ve_252, đừng suy kỳ từ tên."}


# ═════════════════════════════ 3. RAW_ROWS ═════════════════════════════
def _gia_tri_cot(r, cot):
    m = _RE_PAYLOAD.fullmatch(cot)
    if m:
        v = (r.get("_pl") or {}).get(m.group(1))
        return v
    return r.get(cot)


def _kiem_cot(cot):
    if cot in _COT_DB or _RE_PAYLOAD.fullmatch(str(cot or "")):
        return cot
    raise ValueError(f"cột '{cot}' không hỗ trợ. Có: {', '.join(_COT_DB)}, payload.<khoá>")


def _loc_db(loc):
    if loc in (None, "", []):
        return []
    if not isinstance(loc, list):
        raise ValueError("loc phải là danh sách {cot, toan_tu, gia_tri}")
    out = []
    for d in loc:
        if not isinstance(d, dict):
            raise ValueError("mỗi điều kiện loc là object")
        cot = _kiem_cot(d.get("cot"))
        tt = str(d.get("toan_tu") or "=")
        if tt not in ("=", "!=", "in", "chua", "khong_chua"):
            raise ValueError("toan_tu của tl_raw_rows: =, !=, in, chua, khong_chua")
        gt = d.get("gia_tri")
        if tt == "in" and (not isinstance(gt, list) or not gt):
            raise ValueError("toan_tu 'in' cần danh sách")
        out.append((cot, tt, gt))
    return out


def _khop(v, tt, gt):
    if tt in ("=", "!="):
        return tx._bang(v, gt) != (tt == "!=")
    if tt == "in":
        return any(tx._bang(v, g) for g in gt)
    return (tx.chuan(gt) in tx.chuan(v)) != (tt == "khong_chua")


def _doc_raw(env, rt, tu, den):
    rows = sc._ro(env,
                  "SELECT r.id, r.ngay, r.period_month, r.cong_ty, r.khoi, r.cost_center, r.dim1, r.dim2, r.dim3, "
                  "r.amount, r.amount2, r.payload, r.source_file, d.kind AS dataset_kind, d.period AS dataset_period "
                  "FROM raw_rows r LEFT JOIN datasets d ON d.id = r.dataset_id "
                  "WHERE r.report_type = %s AND ((r.ngay >= %s AND r.ngay <= %s) "
                  "OR (r.ngay IS NULL AND r.period_month >= %s AND r.period_month <= %s)) "
                  "ORDER BY r.source_file, r.ngay, r.id LIMIT %s",
                  (rt, tu, den, tu[:7], den[:7], MAX_DONG_DB + 1))
    for r in rows:
        try:
            r["_pl"] = json.loads(r["payload"]) if r.get("payload") else {}
            if not isinstance(r["_pl"], dict):
                r["_pl"] = {}
        except ValueError:
            r["_pl"] = {}
    return rows


def raw_rows(env, report_type, tu, den, nhom_theo=None, loc=None, gioi_han_mau=40):
    """Dòng `raw_rows` của MỘT report_type trong [tu, den] (theo `ngay`; dòng không có ngày thì theo `period_month`).
    Trả: tổng (số dòng, Σamount, Σamount2 — Decimal, không làm tròn), từng source_file (số dòng, tổng, ngày đầu/cuối,
    loại dataset, CỜ `dang_an` nếu file nằm trong hidden_files — builder BỎ các dòng đó), tổng theo `nhom_theo`, ≤40
    dòng mẫu. `loc`/`nhom_theo` nhận cột: source_file, ngay, period_month, cong_ty, khoi, cost_center, dim1..3,
    dataset_kind, payload.<khoá> (vd payload.tinh_trang)."""
    _env(env)
    rt = _rt_hop_le(report_type)
    tu, den = _ngay(tu, "tu"), _ngay(den, "den")
    if tu > den:
        raise ValueError("tu phải ≤ den")
    bl = _loc_db(loc)
    nhom = [] if nhom_theo in (None, "", []) else ([nhom_theo] if isinstance(nhom_theo, str) else list(nhom_theo))
    for c in nhom:
        _kiem_cot(c)
    gioi_han_mau = max(0, min(int(gioi_han_mau or 40), 200))
    rows = _doc_raw(env, rt, tu, den)
    cat = len(rows) > MAX_DONG_DB
    rows = rows[:MAX_DONG_DB]
    an = _an(env)
    rows = [r for r in rows if all(_khop(_gia_tri_cot(r, c), tt, gt) for c, tt, gt in bl)]
    tong = {"so_dong": 0, "amount": Decimal(0), "amount2": Decimal(0)}
    tong_tru_an = {"so_dong": 0, "amount": Decimal(0), "amount2": Decimal(0)}
    theo_sf, theo_nhom = {}, {}
    for r in rows:
        a, a2 = _dec(r["amount"]), _dec(r["amount2"])
        for t in ([tong, tong_tru_an] if r["source_file"] not in an else [tong]):
            t["so_dong"] += 1
            t["amount"] += a
            t["amount2"] += a2
        s = theo_sf.setdefault(r["source_file"], {"so_dong": 0, "amount": Decimal(0), "amount2": Decimal(0),
                                                  "ngay": set(), "ky": set(), "dataset_kind": set()})
        s["so_dong"] += 1
        s["amount"] += a
        s["amount2"] += a2
        if r["ngay"]:
            s["ngay"].add(r["ngay"])
        s["ky"].add(r["period_month"])
        s["dataset_kind"].add(r["dataset_kind"])
        if nhom:
            k = tuple(_gia_tri_cot(r, c) for c in nhom)
            g = theo_nhom.setdefault(k, {"so_dong": 0, "amount": Decimal(0), "amount2": Decimal(0)})
            g["so_dong"] += 1
            g["amount"] += a
            g["amount2"] += a2

    def so3(t):
        return {"so_dong": t["so_dong"], "amount": _so(t["amount"]), "amount2": _so(t["amount2"])}
    ds_sf = [{"source_file": sf, "dang_an": sf in an, **so3(s),
              "ngay_dau": min(s["ngay"]) if s["ngay"] else None, "ngay_cuoi": max(s["ngay"]) if s["ngay"] else None,
              "ky": sorted(k for k in s["ky"] if k), "dataset_kind": sorted(k for k in s["dataset_kind"] if k)}
             for sf, s in sorted(theo_sf.items(), key=lambda kv: -kv[1]["so_dong"])]
    out = {"env": env, "report_type": rt, "tu": tu, "den": den, "loc": loc or [],
           "tong": so3(tong), "tong_bo_file_dang_an": so3(tong_tru_an),
           "theo_source_file": ds_sf, "file_dang_an_co_dong": [x["source_file"] for x in ds_sf if x["dang_an"]],
           "mau": [{k: v for k, v in r.items() if k not in ("payload", "_pl")} | {"payload": r["_pl"]}
                   for r in rows[:gioi_han_mau]],
           "cat_bot_db": cat,
           "lenh": _phu_luc_bang_chung("tl_raw_rows", {"env": env, "report_type": rt, "tu": tu, "den": den,
                                                        "nhom_theo": nhom_theo, "loc": loc}),
           "luu_y": "amount/amount2 là giá trị ĐÃ QUY ĐỔI theo spec (thường amount = tỷ đồng, xem payload.unit). Builder "
                    "bỏ dòng của file trong hidden_files (so khớp ĐÚNG chuỗi source_file) — so với dashboard dùng "
                    "tong_bo_file_dang_an."}
    if nhom:
        ds = sorted(theo_nhom.items(), key=lambda kv: -abs(kv[1]["amount"]) - kv[1]["so_dong"])
        out["theo_nhom"] = [{"nhom": dict(zip(nhom, k)), **so3(g)} for k, g in ds[:200]]
    return out


# ═════════════════════════════ 4. SỐ DASHBOARD ═════════════════════════════
_cache_b = {}
_cache_lock = threading.Lock()


def _env_api(env):
    ev = {}
    with open(os.path.join(mc.API_DIR, ".env"), encoding="utf-8") as fh:
        for ln in fh:
            ln = ln.strip()
            if ln and not ln.startswith("#") and "=" in ln:
                k, _, v = ln.partition("=")
                ev[k.strip()] = v.strip().strip("'\"")
    url = ev.get("DATABASE_URL") or ""
    for p in _PORT.values():
        url = url.replace(f":{p}/", f":{_PORT[env]}/")
    if f":{_PORT[env]}/" not in url:
        raise RuntimeError("không dựng được DATABASE_URL cho env " + env)
    return {**os.environ, **ev, "DATABASE_URL": url, "PYTHONPATH": mc.API_DIR,
            "PGOPTIONS": "-c default_transaction_read_only=on -c statement_timeout=60000"}


def _yeu_cau(man, ky, grain, khoi, cong_ty, tu_ngay, den_ngay):
    if grain not in ("month", "day"):
        raise ValueError("grain phải là month hoặc day")
    if not _RE_MAN.fullmatch(str(man or "")):
        hit = mf.tim_man(man)
        if not hit:
            raise ValueError(f"không nhận ra màn '{man}'")
        man = hit[0]["man"]
    y = {"man": man, "grain": grain, "khoi": khoi or None, "cong_ty": cong_ty or None, "period": None,
         "tu_ngay": None, "den_ngay": None}
    if grain == "month":
        if not re.fullmatch(r"\d{4}-\d{2}", str(ky or "")):
            raise ValueError("grain=month cần ky YYYY-MM")
        y["period"] = ky
    else:
        if tu_ngay or den_ngay:
            y["tu_ngay"], y["den_ngay"] = _ngay(tu_ngay or den_ngay, "tu_ngay"), _ngay(den_ngay or tu_ngay, "den_ngay")
        else:
            y["tu_ngay"], y["den_ngay"] = _ky_sang_khoang(ky)
    for k in ("khoi", "cong_ty"):
        if y[k] is not None and (len(str(y[k])) > 80 or not isinstance(y[k], str)):
            raise ValueError(f"{k} không hợp lệ")
    return y


def _khoa(env, y):
    return (env, y["man"], y["period"], y["grain"], y["khoi"], y["cong_ty"], y["tu_ngay"], y["den_ngay"])


def lay_nhieu(env, ds_y, bo_cache=False, timeout=240):
    """Dựng số dashboard cho nhiều yêu cầu trên MỘT env trong một tiến trình con; cache 5 phút từng yêu cầu.
    -> list (kết quả, lúc_dựng_epoch) cùng thứ tự."""
    _env(env)
    bay = time.time()
    ra, thieu = [None] * len(ds_y), []
    with _cache_lock:
        for i, y in enumerate(ds_y):
            c = _cache_b.get(_khoa(env, y))
            if not bo_cache and c and bay - c[1] < TTL_BUILDER:
                ra[i] = c
            else:
                thieu.append(i)
    if thieu:
        arg = {"yeu_cau": [ds_y[i] for i in thieu]}
        r = subprocess.run([os.path.join(mc.API_DIR, ".venv", "bin", "python"), DUMP, json.dumps(arg, ensure_ascii=False)],
                           capture_output=True, text=True, timeout=timeout, cwd=mc.API_DIR, env=_env_api(env))
        if r.returncode != 0:
            raise RuntimeError("không dựng được số dashboard: " + (r.stderr or "")[-400:])
        kq = json.loads(r.stdout)
        with _cache_lock:
            for i, x in zip(thieu, kq):
                ra[i] = (x, bay)
                if not (isinstance(x, dict) and x.get("loi")):
                    _cache_b[_khoa(env, ds_y[i])] = (x, bay)
            for k in [k for k, v in _cache_b.items() if bay - v[1] >= TTL_BUILDER]:
                _cache_b.pop(k, None)
    return ra


_TRUONG_SO = ("thucHien", "giaTri", "gia_tri", "value", "cuoi", "tong", "soLuong")
_TRUONG_DINH_DANH = ("ma", "id", "key", "code", "name", "ten", "chiTieu", "loaiHang", "nhom", "label")


def lay_key(goc, key):
    """Đi theo đường 'a.b.c' trong JSON builder. Gặp danh sách: đoạn là số -> chỉ số; còn lại -> phần tử có trường định
    danh (ma/id/key/name/ten/chiTieu/loaiHang/nhom/label) bằng đoạn đó (so `chuan`). -> (giá trị, đường đã đi, lỗi)."""
    node, da = goc, []
    for seg in [s for s in re.split(r"\.|\[|\]", str(key or "")) if s != ""]:
        if isinstance(node, dict):
            if seg not in node:
                return None, ".".join(da), f"không có khoá '{seg}' (có: {list(node)[:40]})"
            node = node[seg]
        elif isinstance(node, list):
            if seg.isdigit() and int(seg) < len(node):
                node = node[int(seg)]
            else:
                hit = next((x for x in node if isinstance(x, dict) and any(
                    tx.chuan(x.get(f)) == tx.chuan(seg) for f in _TRUONG_DINH_DANH if x.get(f) is not None)), None)
                if hit is None:
                    mau = [{f: x.get(f) for f in _TRUONG_DINH_DANH if isinstance(x, dict) and f in x} for x in node[:12]]
                    return None, ".".join(da), f"danh sách không có phần tử '{seg}' (mẫu định danh: {mau})"
                node = hit
        else:
            return None, ".".join(da), f"'{'.'.join(da)}' là giá trị đơn, không đi tiếp được '{seg}'"
        da.append(seg)
    return node, ".".join(da), None


def so_cua(node):
    """Số đại diện của một nút: chính nó nếu là số; dict -> trường số đầu tiên trong _TRUONG_SO."""
    if isinstance(node, (int, float)) and not isinstance(node, bool):
        return node
    if isinstance(node, dict):
        for f in _TRUONG_SO:
            v = node.get(f)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return v
    return None


def _khung(node, sau=0):
    """Khung JSON gọn (khi kết quả quá lớn): giữ số/chuỗi ngắn, danh sách cắt 3 phần tử."""
    if sau > 5:
        return "…"
    if isinstance(node, dict):
        return {k: _khung(v, sau + 1) for k, v in list(node.items())[:60]}
    if isinstance(node, list):
        return [_khung(x, sau + 1) for x in node[:3]] + ([f"… +{len(node) - 3} phần tử"] if len(node) > 3 else [])
    if isinstance(node, str) and len(node) > 160:
        return node[:160] + "…"
    return node


def _goc_man(res, man):
    return res.get(man) if isinstance(res, dict) and isinstance(res.get(man), dict) else res


def so_dashboard(env, man, ky=None, grain="month", khoi=None, key=None, cong_ty=None, tu_ngay=None, den_ngay=None,
                 bo_cache=False, gioi_han_ky_tu=60000):
    """Số builder dashboard trả cho (màn, kỳ) — đúng số FE vẽ, quyền admin. Có `key` (vd 'donNhap.tongXe.nhanSl',
    'kpi.sl.thucHien', 'phaiNop.cuoi') -> giá trị đúng key đó + `so`. Không key -> JSON đầy đủ (quá ~60k ký tự thì trả
    khung rút gọn — khi đó hãy hỏi theo key). Cache 5 phút; `bo_cache=true` để dựng lại."""
    y = _yeu_cau(man, ky, grain, khoi, cong_ty, tu_ngay, den_ngay)
    (res, luc), = lay_nhieu(env, [y], bo_cache)
    out = {"env": env, **y, "dung_luc": _gio_vn(luc), "cache_toi_da_giay": TTL_BUILDER,
           "lenh": _phu_luc_bang_chung("tl_so_dashboard", {"env": env, "man": y["man"], "ky": ky, "grain": grain,
                                                            "khoi": khoi, "key": key})}
    if isinstance(res, dict) and res.get("loi"):
        out["loi"] = res["loi"]
        out["ky_co"] = res.get("ky_co")
        return out
    goc = _goc_man(res, y["man"])
    if isinstance(res, dict) and goc is not res:
        out["phan_ngoai_man"] = {k: v for k, v in res.items() if k != y["man"]}
    if key:
        v, duong, loi = lay_key(goc, key)
        out.update({"key": key, "duong_da_di": duong, "gia_tri": v if len(json.dumps(v, default=str)) <= gioi_han_ky_tu
                    else _khung(v), "so": so_cua(v)})
        if loi:
            out["loi_key"] = loi
        return out
    s = json.dumps(goc, ensure_ascii=False, default=str)
    out["ket_qua"] = goc if len(s) <= gioi_han_ky_tu else _khung(goc)
    out["rut_gon"] = len(s) > gioi_han_ky_tu
    return out


# ═════════════════════════════ 5. MÃ NGUỒN ═════════════════════════════
def doc_code(man=None, ham=None, env="prod", so_dong=250):
    """Mã nguồn builder của màn (`man`, qua mapping_logic_dump — file:dòng + report_type) hoặc một hàm bất kỳ trong
    `app/` của tc-admin-api (`ham` = tên hàm). Chỉ đọc. env=prod đọc cây đang chạy prod, env=test đọc cây TEST."""
    _env(env)
    so_dong = max(20, min(int(so_dong or 250), 600))
    if man:
        m = mc._man_hop_le(man)
        out = {"man": man, "builder": m.get("builder"), "file": m.get("file"), "dong_bat_dau": m.get("dong_bat_dau"),
               "so_dong_ham": m.get("so_dong_ham"), "report_types": m.get("report_types"), "rt_ngay": m.get("rt_ngay"),
               "doc": m.get("doc")}
        ham = ham or m.get("builder")
        if not ham:
            return out
    else:
        out = {}
    if not _RE_HAM.fullmatch(str(ham or "")):
        raise ValueError("ham phải là tên hàm Python (chữ/số/_)")
    goc = os.path.realpath(os.path.join(API_DIR[env], "app"))
    rx = re.compile(rf"^(\s*)(?:async\s+)?def\s+{re.escape(ham)}\s*\(")
    tim = []
    for p in sorted(glob.glob(os.path.join(goc, "**", "*.py"), recursive=True)):
        if not os.path.realpath(p).startswith(goc + os.sep):
            continue
        try:
            with open(p, encoding="utf-8") as fh:
                dong = fh.read().splitlines()
        except (OSError, UnicodeDecodeError):
            continue
        for i, ln in enumerate(dong):
            m = rx.match(ln)
            if not m:
                continue
            thut = len(m.group(1))
            j = i + 1
            while j < len(dong) and j - i < so_dong:
                s = dong[j]
                if s.strip() and not s.lstrip().startswith("#") and len(s) - len(s.lstrip()) <= thut \
                        and not s.lstrip().startswith((")", "]", "}")):
                    break
                j += 1
            a = i
            while a > 0 and dong[a - 1].lstrip().startswith("@"):
                a -= 1
            tim.append({"file": os.path.relpath(p, API_DIR[env]), "dong": a + 1,
                        "ma": "\n".join(f"{k + 1}: {dong[k]}" for k in range(a, j)), "cat_bot": j - i >= so_dong})
            if len(tim) >= 5:
                break
    out.update({"env": env, "api_dir": API_DIR[env], "ham": ham, "tim_thay": tim})
    if not tim:
        out["loi"] = f"không thấy 'def {ham}(' trong {API_DIR[env]}/app"
    return out


# ═════════════════════════════ 6. GỘP: HỒ SƠ MÀN ═════════════════════════════
_cache_hs = {}


def _specs_day_du():
    out = []
    for f in sorted(glob.glob(os.path.join(SPEC_DIR, "*.json"))):
        try:
            with open(f, encoding="utf-8") as fh:
                s = json.load(fh)
        except (OSError, ValueError):
            continue
        if isinstance(s, dict) and s.get("report_type"):
            s["_file"] = os.path.basename(f)
            out.append(s)
    return out


def _spec_gon(s):
    ng = s.get("nguon") or {}
    return {"id": s.get("id"), "file_spec": s.get("_file"), "report_type": s.get("report_type"), "khoi": s.get("khoi"),
            "folder": ng.get("folder"), "file_glob": ng.get("file_glob"), "sheet": ng.get("sheet"),
            "bo_qua_file": ng.get("bo_qua_file"), "header": s.get("header"), "dong_bat_dau": s.get("dong_bat_dau"),
            "cot": s.get("cot"), "loc": s.get("loc"), "moi_ky_lay_file_moi_nhat": s.get("moi_ky_lay_file_moi_nhat"),
            "mo_ta": (s.get("_mo_ta") or "")[:1500], "bay": s.get("_bay")}


def _file_cua_spec(s, toi_da=3):
    ng = s.get("nguon") or {}
    folder, pat = ng.get("folder"), ng.get("file_glob") or "*"
    if not folder or ".." in folder.split("/"):
        return []
    d = os.path.join(tx.RECEIVED_DIR, folder)
    if not os.path.isdir(d):
        return []
    bo = ng.get("bo_qua_file")
    fs = [p for p in glob.glob(os.path.join(glob.escape(d), "*"))
          if os.path.isfile(p) and p.lower().endswith(tx.DUOI_EXCEL)
          and fnmatch.fnmatch(os.path.basename(p).lower(), str(pat).lower())
          and not (bo and re.search(bo, os.path.basename(p)))]
    fs.sort(key=os.path.getmtime, reverse=True)
    dv = folder.split("/")[0]
    return [_mo_ta_file(p, dv) for p in fs[:toi_da]] + ([{"con_lai": len(fs) - toi_da}] if len(fs) > toi_da else [])


def _chu_ky_hs():
    def mx(pat):
        return max((os.path.getmtime(p) for p in glob.glob(pat)), default=0)
    ci = mc._chi_muc()
    return (mx(os.path.join(SPEC_DIR, "*.json")), mx(os.path.join(mc.API_DIR, "app", "metrics", "*.py")),
            mx(os.path.join(mf.WEB_SRC, "pages", "*.tsx")), os.path.getmtime(os.path.join(mf.WEB_SRC, "nav.ts")),
            max((os.path.getmtime(os.path.join(ci["thu_muc_tai_lieu"], d["file"])) for d in ci["tai_lieu"].values()
                 if os.path.exists(os.path.join(ci["thu_muc_tai_lieu"], d["file"]))), default=0),
            os.path.getmtime(mc.INDEX))


def _truong_tren_trang(trang, toi_da=120):
    """Mọi trường `d.xxx.yyy` trang FE đọc — đường lùi khi `man_fe.nhan_tren_man` không bắt được nhãn (trang không
    dùng <Kpi label=...>, vd HtxtCungUng.tsx). Key JSON của tl_so_dashboard tương đối với `d` này."""
    if not trang:
        return []
    try:
        with open(trang, encoding="utf-8") as fh:
            s = fh.read()
    except OSError:
        return []
    ra = []
    for m in re.finditer(r"\bd\??\.([A-Za-z_]\w*(?:\??\.[A-Za-z_]\w*)*)", s):
        t = m.group(1).replace("?.", ".")
        if t not in ra:
            ra.append(t)
    return ra[:toi_da]


def ho_so_man(man, tu_khoa=None, bo_cache=False):
    """MỘT lần gọi: builder (file:dòng, docstring), report_type builder đọc (+ nguồn NGÀY), spec trích xuất của từng
    report_type (folder, file_glob, sheet, header, cột→trường, `_bay`), 3 file mới nhất trên đĩa mỗi spec (mtime, lưu
    về 252, có trong available_metadata không), dòng mapping của màn (hoặc khớp `tu_khoa`), nhãn thẻ FE ↔ key JSON.
    Cache theo mtime spec/builder/FE/mapping."""
    t0 = time.time()
    if not _RE_MAN.fullmatch(str(man or "")):
        hit = mf.tim_man(man)
        if not hit:
            raise ValueError(f"không nhận ra màn '{man}'")
        man = hit[0]["man"]
    ck = _chu_ky_hs()
    k = (man, tu_khoa or "")
    if not bo_cache and _cache_hs.get(k, (None,))[0] == ck:
        out = dict(_cache_hs[k][1])
        out["thoi_gian_giay"] = round(time.time() - t0, 3)
        out["tu_cache"] = True
        return out
    m = mc._man_hop_le(man)
    rts = [rt for rt in m.get("report_types") or [] if not re.fullmatch(r"CT\d{2}", rt)]
    specs = _specs_day_du()
    theo_rt = {}
    for s in specs:
        theo_rt.setdefault(s["report_type"], []).append(s)
    spec_ra = {}
    for rt in rts:
        spec_ra[rt] = [{**_spec_gon(s), "file_moi_nhat_tren_dia": _file_cua_spec(s)} for s in theo_rt.get(rt, [])] or \
            "KHÔNG có spec trích xuất (nạp bằng deriver/cron riêng hoặc report_type _D dựng từ nguồn khác)"
    ten = next((x for x in mf.danh_muc_man() if x["man"] == man), None)
    mapping = []
    if tu_khoa:
        mapping = mc.tim_chi_tieu(tu_khoa, man=man, toi_da=12)["ket_qua"]
    else:
        kws = {tx._kd(man)} | ({tx._kd(ten["ten_man"])} if ten else set())
        for tl, d in mc._tai_lieu_cua_man(man):
            for sh in d["sheet"][:1]:                        # sheet đầu = bản mới nhất đã khai
                try:
                    rows = mc._doc_sheet(d["file"], sh)
                except (OSError, ValueError):
                    continue
                khop = [r for r in rows if any(w and w in tx._kd(" ".join(
                    str(r.get(f) or "") for f in ("man_hinh", "nhom", "nhom_chi_tieu"))) for w in kws)]
                # Tài liệu không có cột 'Màn hình' (HTXT, DUAN…) mà chỉ phục vụ ≤3 màn: đưa cả sheet (gọn) — bot tự
                # chọn dòng, đỡ một lượt tl_doc_mapping dò mù. Tài liệu nhiều màn (QTTC 12 màn) thì phải có tu_khoa.
                if not khop and len(d["man"]) <= 3:
                    khop = rows
                mapping += [{"tai_lieu": tl, "file": d["file"], "sheet": sh,
                             **{f: r.get(f) for f in ("dong", "nhom", "nhom_chi_tieu", "chi_tieu", "ma", "cong_thuc",
                                                      "nguon")}} for r in khop]
    gon_map = [{k2: (str(v)[:300] if isinstance(v, str) else v) for k2, v in r.items() if v not in (None, "")}
               for r in mapping[:60]]
    trang = mf._file_trang(man)
    out = {"man": man, "ten_man": ten and ten["ten_man"], "nhom_man": ten and ten["nhom"],
           "builder": {"ham": m.get("builder"), "file": f"{m.get('file')}:{m.get('dong_bat_dau')}",
                       "so_dong_ham": m.get("so_dong_ham"), "doc": (m.get("doc") or "")[:1200]},
           "report_types": rts, "rt_ngay": m.get("rt_ngay"), "spec_theo_report_type": spec_ra,
           "tai_lieu_mapping": [{"tai_lieu": k2, "file": d["file"], "sheet": d["sheet"],
                                 "ban_cung_ten": _ban_copy(d["file"])} for k2, d in mc._tai_lieu_cua_man(man)],
           "dong_mapping": gon_map, "so_dong_mapping": len(mapping), "cat_bot_mapping": max(0, len(mapping) - 60),
           "fe": {"file_trang": trang and os.path.relpath(trang, os.path.dirname(mf.WEB_SRC)),
                  "nhan_the": mf.nhan_tren_man(man), "truong_d_tren_trang": _truong_tren_trang(trang)},
           "luu_y": "Hồ sơ tĩnh (code/spec/file/mapping). Số thì gọi tl_so_dashboard / tl_truy_vet; nạp thì tl_kiem_nhanh_nap."}
    _cache_hs[k] = (ck, out)
    out = dict(out)
    out["thoi_gian_giay"] = round(time.time() - t0, 3)
    return out


# ═════════════════════════════ 7. GỘP: KIỂM NHANH TẦNG NẠP ═════════════════════════════
def _log_cron(env, ids, doc_toi=600_000):
    """Dòng log cron GẦN NHẤT nhắc tới mỗi spec id (logs ở AI_coding/logs; prod = *_prod.log, test = phần còn lại trừ
    dryrun). -> {id: {log, job, dong, luc}}."""
    ra = {}
    rx = {i: re.compile(rf"(?<![\w]){re.escape(i)}(?![\w])") for i in ids}
    for p in sorted(glob.glob(os.path.join(LOGS_DIR, "cron_*.log"))):
        b = os.path.basename(p)
        if "dryrun" in b or (env == "prod") != b.endswith("_prod.log"):
            continue
        try:
            with open(p, "rb") as fh:
                fh.seek(0, 2)
                n = fh.tell()
                fh.seek(max(0, n - doc_toi))
                txt = fh.read().decode("utf-8", "replace").splitlines()
        except OSError:
            continue
        for ln in reversed(txt):
            for i, r in rx.items():
                if r.search(ln):
                    m = re.match(r"^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", ln)
                    luc = m.group(1) if m else None
                    if i not in ra or (luc and (ra[i]["luc"] or "") < luc):
                        ra[i] = {"log": b, "job": re.sub(r"^cron_|(_prod)?\.log$", "", b), "dong": ln[:400], "luc": luc}
    return ra


def kiem_nhanh_nap(doi_tuong, ky=None, env="prod"):
    """Bẫy tầng NẠP trong một lần, cho một report_type (HTXT_GIAO) hoặc một folder nguồn
    (HUNGTHINH/baocaocapnhatvanhanhxetai). `ky` YYYY-MM | YYYY-MM-DD (bỏ trống = mọi kỳ). Trả từng source_file (đếm,
    tổng, kỳ) và `canh_bao[]` có mã:
      HIDDEN            file đang có dòng mà nằm trong hidden_files -> builder bỏ, dashboard trống/thiếu
      CONG_DOI          ≥2 source_file CÙNG đơn vị phủ CÙNG ngày/kỳ -> nghi cộng đôi / phát hành lại tên khác
      FILE_MOI_CHUA_NAP file mới nhất trên đĩa lưu về SAU lượt nạp gần nhất, hoặc chưa có dòng nào trong DB
      CRON_LOI          dòng log cron gần nhất của spec báo lỗi (rc≠0 / Traceback / LỖI / DỪNG)
      KHONG_CO_CRON     không thấy dòng log cron nào nhắc spec (nạp tay?) — chỉ là thông tin."""
    t0 = time.time()
    _env(env)
    specs = _specs_day_du()
    s_dt = str(doi_tuong or "").strip()
    if _RE_RT.fullmatch(s_dt):
        rts = [s_dt]
        sp = [s for s in specs if s["report_type"] == s_dt]
    else:
        if ".." in s_dt.split("/") or not s_dt:
            raise ValueError("doi_tuong là report_type (HTXT_GIAO) hoặc folder nguồn (ĐƠNVỊ/thưmục)")
        sp = [s for s in specs if str((s.get("nguon") or {}).get("folder") or "").lower() == s_dt.lower()]
        if not sp:
            raise ValueError(f"không spec nào đọc folder '{s_dt}'")
        rts = sorted({s["report_type"] for s in sp})
    if ky:
        tu, den = _ky_sang_khoang(ky)
    an = _an(env)
    canh_bao, theo_rt = [], {}
    for rt in rts:
        q = ("SELECT source_file, COALESCE(ngay, period_month) AS k, COUNT(*) AS n, SUM(amount) AS a, SUM(amount2) AS a2 "
             "FROM raw_rows WHERE report_type=%s")
        p = [rt]
        if ky:
            q += " AND ((ngay >= %s AND ngay <= %s) OR (ngay IS NULL AND period_month >= %s AND period_month <= %s))"
            p += [tu, den, tu[:7], den[:7]]
        rows = sc._ro(env, q + " GROUP BY 1, 2", tuple(p))
        sf = {}
        for r in rows:
            x = sf.setdefault(r["source_file"], {"so_dong": 0, "amount": Decimal(0), "amount2": Decimal(0), "k": set()})
            x["so_dong"] += r["n"]
            x["amount"] += _dec(r["a"])
            x["amount2"] += _dec(r["a2"])
            x["k"].add(r["k"])
        theo_rt[rt] = [{"source_file": f, "dang_an": f in an, "so_dong": x["so_dong"], "amount": _so(x["amount"]),
                        "amount2": _so(x["amount2"]), "ky_dau": min(k for k in x["k"] if k) if any(x["k"]) else None,
                        "ky_cuoi": max(k for k in x["k"] if k) if any(x["k"]) else None, "so_ngay_ky": len(x["k"])}
                       for f, x in sorted(sf.items(), key=lambda kv: str(kv[0]))]
        for f in sf:
            if f in an:
                canh_bao.append({"ma": "HIDDEN", "report_type": rt, "source_file": f, "so_dong": sf[f]["so_dong"],
                                 "y_nghia": "file nằm trong hidden_files — builder bỏ toàn bộ dòng của file này"})
        ds = sorted(sf.items(), key=lambda kv: str(kv[0]))
        for i in range(len(ds)):
            for j in range(i + 1, len(ds)):
                (f1, x1), (f2, x2) = ds[i], ds[j]
                if str(f1).split("::")[0] != str(f2).split("::")[0] or f1 in an or f2 in an:
                    continue
                chung = sorted(k for k in (x1["k"] & x2["k"]) if k)
                if chung:
                    canh_bao.append({"ma": "CONG_DOI", "report_type": rt, "source_file": [f1, f2],
                                     "ngay_ky_trung": chung[:10], "so_ngay_ky_trung": len(chung),
                                     "y_nghia": "hai file cùng đơn vị cùng phủ một ngày/kỳ — nghi cộng đôi hoặc phát "
                                                "hành lại tên khác (có thể là cố ý nếu hai file là hai phần bổ sung "
                                                "nhau: kiểm dim/cost_center bằng tl_raw_rows nhom_theo)"})
    # spec: file đĩa vs lượt nạp gần nhất
    nhat_ky = _log_cron(env, [s["id"] for s in sp if s.get("id")])
    spec_ra = []
    for s in sp:
        fs = [f for f in _file_cua_spec(s, toi_da=1) if "ten" in f]
        lg = nhat_ky.get(s.get("id"))
        moi = fs[0] if fs else None
        x = {"spec": s.get("id"), "report_type": s["report_type"], "folder": (s.get("nguon") or {}).get("folder"),
             "file_moi_nhat": moi, "cron_gan_nhat": lg}
        spec_ra.append(x)
        if not lg:
            canh_bao.append({"ma": "KHONG_CO_CRON", "spec": s.get("id"),
                             "y_nghia": "không thấy log cron nào nhắc spec này — nguồn nạp tay hoặc job chưa đăng ký"})
        elif _RE_LOI_CRON.search(lg["dong"]):
            canh_bao.append({"ma": "CRON_LOI", "spec": s.get("id"), "dong_log": lg["dong"], "log": lg["log"]})
        if moi:
            luu = moi.get("luu_ve_252") or moi.get("sua_tren_252")
            if lg and lg.get("luc") and luu and luu > lg["luc"]:
                canh_bao.append({"ma": "FILE_MOI_CHUA_NAP", "spec": s.get("id"), "file": moi["ten"], "luu_ve_252": luu,
                                 "lan_nap_gan_nhat": lg["luc"], "y_nghia": "file về SAU lượt nạp gần nhất"})
            ten_db = {str(f).split("::", 1)[-1] for f in (r["source_file"] for r in theo_rt.get(s["report_type"], []))}
            if not ky and moi["ten"] not in ten_db:
                canh_bao.append({"ma": "FILE_MOI_CHUA_NAP", "spec": s.get("id"), "file": moi["ten"],
                                 "y_nghia": "file mới nhất trên đĩa chưa có dòng nào trong DB (chưa nạp, bị bộ lọc spec "
                                            "loại, hoặc spec chỉ lấy một file khác)"})
    return {"env": env, "doi_tuong": s_dt, "ky": ky, "report_types": rts, "theo_report_type": theo_rt, "spec": spec_ra,
            "canh_bao": canh_bao, "ma_canh_bao": sorted({c["ma"] for c in canh_bao}),
            "thoi_gian_giay": round(time.time() - t0, 3),
            "luu_y": "Mã cảnh báo là CỜ để đi kiểm tiếp, không phải kết luận. Giờ log và giờ file đều là giờ VN."}


# ═════════════════════════════ 8. GỘP: MỘT KEY QUA NHIỀU KỲ ═════════════════════════════
def so_qua_ky(env, man, key, cac_ky, grain="month", khoi=None, cong_ty=None, bo_cache=False):
    """Đo MỘT key qua nhiều kỳ (tối đa 24) — env 'prod' | 'test' | 'ca_hai' (đo cả hai DB, song song) -> bảng
    {ky, env, so, gia_tri}. Dùng cho 'kiểm các tháng khác'. Cache 5 phút như tl_so_dashboard."""
    t0 = time.time()
    if not isinstance(cac_ky, list) or not cac_ky or len(cac_ky) > 24:
        raise ValueError("cac_ky là danh sách 1..24 kỳ")
    envs = ["prod", "test"] if env == "ca_hai" else [_env(env)]
    ds_y = [_yeu_cau(man, k, grain, khoi, cong_ty, None, None) for k in cac_ky]
    with ThreadPoolExecutor(max_workers=2) as ex:
        kq = dict(zip(envs, ex.map(lambda e: lay_nhieu(e, ds_y, bo_cache), envs)))
    bang = []
    for i, k in enumerate(cac_ky):
        for e in envs:
            res, luc = kq[e][i]
            if isinstance(res, dict) and res.get("loi"):
                bang.append({"ky": k, "env": e, "so": None, "loi": res["loi"]})
                continue
            v, _, loi = lay_key(_goc_man(res, ds_y[i]["man"]), key)
            dong = {"ky": k, "env": e, "so": so_cua(v), "dung_luc": _gio_vn(luc)}
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                dong["gia_tri"] = _khung(v)
            if loi:
                dong["loi_key"] = loi
            bang.append(dong)
    out = {"man": ds_y[0]["man"], "key": key, "grain": grain, "khoi": khoi, "bang": bang,
           "thoi_gian_giay": round(time.time() - t0, 3), "cache_toi_da_giay": TTL_BUILDER,
           "lenh": _phu_luc_bang_chung("tl_so_qua_ky", {"env": env, "man": man, "key": key, "cac_ky": cac_ky,
                                                         "grain": grain, "khoi": khoi})}
    if env == "ca_hai":
        out["lech_prod_test"] = [k for k in cac_ky
                                 if len({json.dumps(d["so"]) for d in bang if d["ky"] == k}) > 1]
    return out


# ═════════════════════════════ 9. GỘP: TRUY VẾT ═════════════════════════════
def _gan(a, b):
    if a is None or b is None:
        return False
    a, b = float(a), float(b)
    return abs(a - b) <= max(1e-9, 1e-9 * max(abs(a), abs(b)))


def truy_vet(man, key, ky, env="prod", so_dung=None, grain="month", khoi=None, bo_cache=False):
    """Chuỗi file → raw_rows → builder → key trong MỘT lần: số dashboard của key; report_type builder đọc (từ hồ sơ
    màn) và với từng report_type: tổng raw_rows trong kỳ + từng source_file + cờ hidden + file mới nhất trên đĩa. Đưa
    `so_dung` (số đúng theo mapping, CÙNG đơn vị với key) thì kèm `goi_y_tang_lech` — là GỢI Ý dựa trên phép so số,
    không phải kết luận: tổng raw_rows thô hiếm khi bằng đúng key (builder còn lọc/nhóm)."""
    t0 = time.time()
    _env(env)
    tu, den = _ky_sang_khoang(ky)
    d = so_dashboard(env, man, ky, grain, khoi, key, bo_cache=bo_cache)
    hs = ho_so_man(d["man"], bo_cache=bo_cache)
    an = _an(env)
    tang = []
    for rt in hs["report_types"][:15]:
        try:
            rows = _doc_raw(env, rt, tu, den)
        except Exception as e:  # noqa: BLE001
            tang.append({"report_type": rt, "loi": str(e)[:200]})
            continue
        sf = {}
        for r in rows:
            x = sf.setdefault(r["source_file"], [0, Decimal(0), Decimal(0)])
            x[0] += 1
            x[1] += _dec(r["amount"])
            x[2] += _dec(r["amount2"])
        bo_an = [x for f, x in sf.items() if f not in an]
        sp = hs["spec_theo_report_type"].get(rt)
        tang.append({"report_type": rt, "so_dong": sum(x[0] for x in sf.values()),
                     "amount": _so(sum((x[1] for x in sf.values()), Decimal(0))),
                     "amount2": _so(sum((x[2] for x in sf.values()), Decimal(0))),
                     "amount_bo_file_an": _so(sum((x[1] for x in bo_an), Decimal(0))),
                     "amount2_bo_file_an": _so(sum((x[2] for x in bo_an), Decimal(0))),
                     "source_file": [{"source_file": f, "so_dong": x[0], "amount": _so(x[1]), "amount2": _so(x[2]),
                                      "dang_an": f in an} for f, x in sorted(sf.items(), key=lambda kv: -kv[1][0])][:8],
                     "spec": [s["id"] for s in sp] if isinstance(sp, list) else sp,
                     "file_moi_nhat_tren_dia": [f.get("ten") for s in sp for f in s["file_moi_nhat_tren_dia"]
                                                if f.get("ten")][:3] if isinstance(sp, list) else []})
    out = {"man": d["man"], "key": key, "ky": ky, "env": env,
           "1_dashboard": {k: d.get(k) for k in ("so", "gia_tri", "duong_da_di", "loi_key", "loi", "dung_luc")},
           "2_builder": hs["builder"], "3_raw_rows_theo_report_type": tang,
           "4_fe": [n for n in hs["fe"]["nhan_the"] if str(key).split(".")[0] in n["truong"]][:10],
           "thoi_gian_giay": None,
           "lenh": _phu_luc_bang_chung("tl_truy_vet", {"man": man, "key": key, "ky": ky, "env": env, "so_dung": so_dung,
                                                        "grain": grain, "khoi": khoi})}
    if so_dung is not None:
        sd = float(so_dung)
        sdb = d.get("so")
        if _gan(sdb, sd):
            g = {"tang": "khong", "ly_do": f"dashboard {sdb} = số đúng {sd}"}
        else:
            khop_raw = [(t["report_type"], f) for t in tang for f in ("amount_bo_file_an", "amount2_bo_file_an")
                        if _gan(t.get(f), sd) or _gan(t.get(f), sd / 1e9)]
            khop_an = [(t["report_type"], f) for t in tang for f in ("amount", "amount2")
                       if (_gan(t.get(f), sd) or _gan(t.get(f), sd / 1e9))
                       and not _gan(t.get(f), t.get(f + "_bo_file_an"))]
            if khop_raw:
                g = {"tang": "builder", "ly_do": f"raw_rows {khop_raw} đã bằng số đúng mà key ra {sdb} — nghi builder/"
                                                 "key lọc/nhóm lệch mapping"}
            elif khop_an:
                g = {"tang": "nap", "ly_do": f"raw_rows chỉ bằng số đúng khi TÍNH CẢ file đang ẩn {khop_an} — "
                                             "file bị hidden_files che"}
            elif all(t.get("so_dong", 0) == 0 for t in tang if "loi" not in t):
                g = {"tang": "nap", "ly_do": "không report_type nào của màn có dòng trong kỳ — chưa nạp, hoặc FILE chưa "
                                             "về (kiểm tl_kiem_nhanh_nap / tl_tim_file_nguon)"}
            else:
                g = {"tang": "chua_xac_dinh",
                     "ly_do": "tổng thô không khớp số đúng — cần tl_raw_rows có loc/nhom_theo đúng điều kiện builder "
                              "và tl_tong_hop_excel trên file nguồn để biết lệch ở spec hay builder"}
        g["so_dung"], g["so_dashboard"] = sd, sdb
        out["goi_y_tang_lech"] = g
    out["thoi_gian_giay"] = round(time.time() - t0, 3)
    return out

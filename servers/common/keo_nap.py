# -*- coding: utf-8 -*-
"""KÉO FILE NGUỒN + NẠP DB cho bot `ai_keo_file_bot` (14/10/2026) — làm đúng việc user vẫn làm tay:
`pull_nguon.keo` (kéo file từ máy kế toán về 252) rồi `agent_cli autofill` vào TEST, đối chiếu số trước/sau,
rồi mới PROD.

THIẾT KẾ AN TOÀN (user chốt 14/10/2026: giống bot kiểm số liệu, riêng lên PROD cần xác nhận của người nhắn):
  1. KẾ HOẠCH TRƯỚC: `lap_ke_hoach` chỉ nhận tên file CÓ trong `available_metadata.json` (không đường dẫn tự do),
     tối đa MAX_FILES file/lượt, kỳ suy từ tên file (lùi về year/month metadata). Mọi bước sau chỉ chạy theo `plan_id`.
  2. TEST TRƯỚC, PROD SAU: `nap(env="prod")` bị từ chối nếu kế hoạch chưa nạp TEST thành công, hoặc file đã đổi
     (sha1) kể từ lần nạp TEST.
  3. XÁC NHẬN CỦA NGƯỜI NHẮN KHÔNG DO LLM TỰ TẠO: API (`bot_service`) ghi file `xacnhan/<MÃ>.json` khi và chỉ khi
     tin của NGƯỜI DÙNG chứa `XAC NHAN <MÃ>`. Tool đọc file đó — LLM không có đường nào ghi vào thư mục này.
     Mã là của riêng từng kế hoạch, dùng một lần, hết hạn sau XN_TTL_PHUT, và phải đến SAU khi kế hoạch được lập.
     Đợt 1 an toàn (15/10/2026, hợp đồng `wt/hopdong_dot1_an_toan_bot_20261015.md` mục 4): API ghi thêm
     `ma_cua/<MÃ>.json` (chủ mã = người gửi tin kích hoạt lượt bot đầu tiên nhắc mã) → xác nhận PROD và ảnh duyệt
     năm cũ chỉ hợp lệ khi `user_id` trùng chủ mã (người khác trong kênh gõ hộ cũng không qua). Mã bị đánh dấu
     `da_dung` NGAY TRƯỚC khi nạp PROD — nạp hỏng giữa chừng cũng không dùng lại được.
  4. Khoá chạy một lượt một lúc (flock) + khoá từng file của `filelock` bên trong `agent_cli` chống nạp chồng cron.
  5. Nhật ký mọi thao tác ghi: `audit.jsonl`.
Không tool nào nhận SQL hay lệnh tự do.
"""
import datetime
import fcntl
import hashlib
import json
import os
import re
import secrets
import subprocess
import sys
import time

import psycopg
from psycopg.rows import dict_row

ROOT = os.environ.get("TC_ROOT") or "/home/itadmin/AI_Dashboard_QT"
AGENT_DIR = os.environ.get("TC_AGENT_DIR") or os.path.join(ROOT, "Dashboard_Agent")
AGENT_PY = os.path.join(AGENT_DIR, ".venv", "bin", "python")
RECEIVED_DIR = os.environ.get("RECEIVED_DIR") or os.path.join(ROOT, "Connect_VPS", "received_reports")
STATE_DIR = os.environ.get("PULL_BOT_DIR") or os.path.join(AGENT_DIR, "memory", "pull_bot")
PLAN_DIR = os.path.join(STATE_DIR, "plans")
XN_DIR = os.environ.get("PULL_XAC_NHAN_DIR") or os.path.join(STATE_DIR, "xacnhan")
MA_CUA_DIR = os.environ.get("PULL_MA_CUA_DIR") or os.path.join(STATE_DIR, "ma_cua")          # API ghi, tool chỉ đọc
ANH_DUYET_DIR = os.environ.get("PULL_ANH_DUYET_DIR") or os.path.join(STATE_DIR, "anh_duyet")  # API ghi, tool chỉ đọc
AUDIT = os.path.join(STATE_DIR, "audit.jsonl")
LOCK = os.path.join(STATE_DIR, "run.lock")

MAX_FILES = 40
XN_TTL_PHUT = 30
PLAN_TTL_GIO = 6
_PORT = {"prod": 5434, "test": 5435}
_RE_NGAY = re.compile(r"\.[DM]\.(\d{4})(\d{2})(\d{2})(?!\d)")
_RE_TEN = re.compile(r"[\w .()\-]+\.(xlsx|xlsm|xlsb|xls)", re.IGNORECASE)


# ───────────────────────── tiện ích ─────────────────────────
def _bay_gio():
    return datetime.datetime.now()


def _mk():
    for d in (PLAN_DIR, XN_DIR):
        os.makedirs(d, exist_ok=True)


def _doc_json(path):
    """dict trong file JSON, None nếu thiếu/hỏng/không phải dict (file do API ghi — không tin cấu trúc)."""
    try:
        with open(path, encoding="utf-8") as fh:
            x = json.load(fh)
    except (OSError, ValueError):
        return None
    return x if isinstance(x, dict) else None


def _ma_hop_le(ma):
    return bool(re.fullmatch(r"[A-Z0-9]{6}", str(ma or "")))


def _chu_ma(ma):
    """`ma_cua/<MÃ>.json` do API ghi: ai là chủ mã (người đã nhắn khi bot lần đầu đưa mã). None nếu chưa có."""
    if not _ma_hop_le(ma):
        return None
    x = _doc_json(os.path.join(MA_CUA_DIR, ma + ".json"))
    return x if x and x.get("user_id") is not None else None


def _audit(hanh_dong, ma=None, **kv):
    """Một dòng audit.jsonl. Có `ma` → thêm `user_id` của chủ mã (None khi API chưa ghi ma_cua)."""
    _mk()
    if ma is not None:
        kv["user_id"] = (_chu_ma(ma) or {}).get("user_id")
    with open(AUDIT, "a", encoding="utf-8") as fh:
        fh.write(json.dumps({"luc": _bay_gio().isoformat(timespec="seconds"), "hanh_dong": hanh_dong, **kv},
                            ensure_ascii=False) + "\n")


def _doc_meta():
    sys.path.insert(0, os.path.join(AGENT_DIR, "cron"))
    import pull_nguon  # noqa: E402
    return pull_nguon, pull_nguon.read_metadata() or []


def _ky_cua_file(e):
    """'YYYY-MM' từ tên file/metadata, None nếu không suy được (file đó bị bỏ, không đoán).

    Thứ tự: (1) tên `.D.YYYYMMDD`/`.M.YYYYMM` — LUÔN ưu tiên (metadata của SRVF có 56 file lệch tên, vd bản
    phát hành lại giữ tháng cũ); (2) lùi về `year`/`month` của metadata do máy kế toán quét — cách
    `cron_thuchi_daily.pick_targets` vẫn dùng. Thiếu (2) thì 15/15 file THUCHI/baocaothuchi
    ('01_TC01_BÁO CÁO TIỀN TẬP ĐOÀN_THÁNG 01 NĂM 2026.xlsx') và 50/50 baocaonganhang ra None -> bot từ chối
    kéo (15/10/2026). Chặn vắt năm như cron: tên có số 20xx mà không chứa `year` thì không tin metadata."""
    m = _RE_NGAY.search(e["fileName"])
    if m and 1 <= int(m[2]) <= 12:
        return f"{m[1]}-{m[2]}"
    m = re.search(r"\.M\.(20\d{2})[._]?(\d{1,2})(?!\d)", e["fileName"])
    if m and 1 <= int(m[2]) <= 12:
        return f"{m[1]}-{int(m[2]):02d}"
    try:
        nam, thang = int(e.get("year")), int(e.get("month"))
    except (TypeError, ValueError):
        return None
    if not (2000 <= nam <= 2099 and 1 <= thang <= 12):
        return None
    nam_ten = {int(t) for t in re.findall(r"(?<!\d)(20\d{2})(?!\d)", e["fileName"])}
    if nam_ten and nam not in nam_ten:
        return None
    return f"{nam}-{thang:02d}"


def _duong_dan(f):
    return os.path.join(RECEIVED_DIR, f["company"], f["report_type"], f["fileName"])


def _sha1(p):
    h = hashlib.sha1()
    with open(p, "rb") as fh:
        for kh in iter(lambda: fh.read(1 << 20), b""):
            h.update(kh)
    return h.hexdigest()


def _plan_path(pid):
    if not re.fullmatch(r"[0-9a-f]{8}", str(pid or "")):
        raise ValueError("plan_id không hợp lệ")
    return os.path.join(PLAN_DIR, pid + ".json")


def _doc_plan(pid):
    try:
        with open(_plan_path(pid), encoding="utf-8") as fh:
            p = json.load(fh)
    except OSError:
        raise ValueError(f"không có kế hoạch {pid}") from None
    han = datetime.datetime.fromisoformat(p["tao_luc"]) + datetime.timedelta(hours=PLAN_TTL_GIO)
    if _bay_gio() > han:
        raise ValueError(f"kế hoạch {pid} đã hết hạn ({PLAN_TTL_GIO} giờ) — lập lại")
    return p


def _luu_plan(p):
    _mk()
    tmp = _plan_path(p["plan_id"]) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(p, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, _plan_path(p["plan_id"]))


class _KhoaChay:
    """Mỗi lúc một lượt kéo/nạp. Không chờ: có lượt khác đang chạy thì báo bận."""

    def __enter__(self):
        _mk()
        self.fh = open(LOCK, "w")
        try:
            fcntl.flock(self.fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self.fh.close()
            raise RuntimeError("đang có một lượt kéo/nạp khác chạy — thử lại sau ít phút") from None
        return self

    def __exit__(self, *a):
        fcntl.flock(self.fh, fcntl.LOCK_UN)
        self.fh.close()


def _db(env):
    from urllib.parse import unquote, urlparse
    dsn = os.environ.get("DATABASE_URL") or ""
    mk = unquote(urlparse(dsn).password or "") if dsn else ""
    return psycopg.connect(host="localhost", port=_PORT[env], dbname="tc_dashboard", user="tc",
                           password=mk or "tc_$production", autocommit=True, row_factory=dict_row,
                           prepare_threshold=None, options="-c statement_timeout=60000")


def _anh_chup(env, company, periods):
    """{(report_type, ngay): (số dòng, tổng amount)} của đơn vị trong các kỳ — để so trước/sau."""
    with _db(env) as c:
        rows = c.execute(
            "SELECT report_type, ngay, COUNT(*) AS n, COALESCE(SUM(amount),0) AS s FROM raw_rows "
            "WHERE source_file LIKE %s AND period_month = ANY(%s) GROUP BY report_type, ngay",
            (company + "::%", sorted(periods))).fetchall()
    return {(r["report_type"], r["ngay"]): (r["n"], round(float(r["s"]), 6)) for r in rows}


def _so_sanh(truoc, sau):
    doi = []
    for k in sorted(set(truoc) | set(sau), key=lambda x: (x[0], x[1] or "")):
        a, b = truoc.get(k), sau.get(k)
        if a != b:
            doi.append({"report_type": k[0], "ngay": k[1], "truoc": a, "sau": b,
                        "chenh_tong": None if not (a and b) else round(b[1] - a[1], 6)})
    return {"so_nhom_doi": len(doi), "so_nhom": len(set(truoc) | set(sau)), "chi_tiet": doi[:25],
            "cat_bot": max(0, len(doi) - 25)}


# ───────────────────────── 1. KẾ HOẠCH ─────────────────────────
def _ky_nam_cu(ky):
    """Kỳ 'YYYY-MM' thuộc năm TRƯỚC năm hiện tại."""
    return int(ky[:4]) < _bay_gio().year


def lap_ke_hoach(don_vi, thu_muc, ten_file=None, ngay_tu=None, ngay_den=None, period=None, da_co_anh_duyet=False):
    """Chọn file theo: danh sách `ten_file` · hoặc khoảng `ngay_tu..ngay_den` (YYYY-MM-DD, ngày trong TÊN file) ·
    hoặc `period` (YYYY-MM). Chỉ file có trong danh sách nguồn; tối đa MAX_FILES.

    KỲ NĂM CŨ (luật Trưởng phòng KSNB Mai Thị Điệp 15/10/2026: "YC IT kéo lại số các NĂM CŨ phải đính kèm ảnh
    email chị duyệt"): vẫn LẬP được kế hoạch, nhưng plan ghi `can_anh_duyet=[kỳ…]` và `nap` (cả TEST lẫn PROD)
    chỉ chạy khi có `anh_duyet/<MÃ>.json` do API ghi (tin `DUYET <MÃ>` kèm ảnh, đúng chủ mã). Cờ cũ
    `da_co_anh_duyet` do LLM tự đặt (bản 15/10 sáng) — LLM nói dối được — nay nhận nhưng BỎ QUA, ghi audit
    `bo_qua_co_llm` để hậu kiểm."""
    if not (ten_file or ngay_tu or ngay_den or period):
        raise ValueError("phải nêu ten_file, hoặc ngay_tu/ngay_den, hoặc period — không kéo 'tất cả'")
    _pn, meta = _doc_meta()
    cd = [e for e in meta if e.get("fileName") and str(e.get("company") or "").upper() == str(don_vi or "").upper()
          and str(e.get("report_type") or "").lower() == str(thu_muc or "").lower()]
    if not cd:
        co = sorted({(e.get("company"), e.get("report_type")) for e in meta if e.get("company")})
        raise ValueError(f"không có nguồn {don_vi}/{thu_muc} trong danh sách nguồn (có: {co[:40]})")
    chon, bo = [], []
    ten = [ten_file] if isinstance(ten_file, str) else (list(ten_file) if ten_file else [])
    for t in ten:
        if not _RE_TEN.fullmatch(os.path.basename(str(t))):
            raise ValueError(f"tên file không hợp lệ: {t!r}")
    ten_set = {os.path.basename(t).lower() for t in ten}
    for e in cd:
        if ten_set:
            if e["fileName"].lower() not in ten_set:
                continue
        else:
            ky = _ky_cua_file(e)
            m = _RE_NGAY.search(e["fileName"])
            ng = f"{m[1]}-{m[2]}-{m[3]}" if m else None
            if period and ky != period:
                continue
            if ngay_tu or ngay_den:
                if not ng or (ngay_tu and ng < ngay_tu) or (ngay_den and ng > ngay_den):
                    continue
        ky = _ky_cua_file(e)
        if not ky:
            bo.append({"file": e["fileName"], "ly_do": "không suy được kỳ từ tên file/metadata"})
            continue
        chon.append({"company": e["company"], "report_type": e["report_type"], "fileName": e["fileName"],
                     "period": ky})
    thieu = sorted(ten_set - {f["fileName"].lower() for f in chon})
    ky_cu = sorted({f["period"] for f in chon if _ky_nam_cu(f["period"])})
    if len(chon) > MAX_FILES:
        raise ValueError(f"khớp {len(chon)} file, vượt trần {MAX_FILES}/lượt — thu hẹp khoảng ngày")
    if not chon:
        raise ValueError("không có file nào khớp điều kiện trong danh sách nguồn")
    pid = secrets.token_hex(4)
    ma = "".join(secrets.choice("ABCDEFGHJKLMNPQRSTUVWXYZ23456789") for _ in range(6))
    p = {"plan_id": pid, "ma_xac_nhan": ma, "tao_luc": _bay_gio().isoformat(timespec="seconds"),
         "don_vi": chon[0]["company"], "thu_muc": chon[0]["report_type"], "files": sorted(chon, key=lambda f: f["fileName"]),
         "trang_thai": {"keo": None, "test": None, "prod": None}, "can_anh_duyet": ky_cu or None}
    _luu_plan(p)
    if da_co_anh_duyet:
        _audit("bo_qua_co_llm", ma=ma, plan=pid, co="da_co_anh_duyet", gia_tri=bool(da_co_anh_duyet))
    _audit("lap_ke_hoach", ma=ma, plan=pid, don_vi=p["don_vi"], thu_muc=p["thu_muc"], so_file=len(chon),
           **({"can_anh_duyet": ky_cu} if ky_cu else {}))
    kq = {"plan_id": pid, "ma_xac_nhan": ma, "don_vi": p["don_vi"], "thu_muc": p["thu_muc"],
          "so_file": len(chon), "ky": sorted({f["period"] for f in chon}),
          "file": [f["fileName"] for f in chon], "khong_thay_trong_nguon": thieu, "bo_qua": bo,
          "buoc_tiep": "pull_keo_ve -> pull_nap(env='test') -> người dùng gõ 'XAC NHAN <ma_xac_nhan>' -> pull_nap(env='prod')"}
    if ky_cu:
        kq["can_anh_duyet"] = ky_cu
        kq["huong_dan_cho_nguoi_dung"] = (f"Kỳ năm cũ ({', '.join(ky_cu)}) cần ảnh email duyệt của KSNB: gửi tin "
                                          f"`DUYET {ma}` kèm ảnh")
        kq["buoc_tiep"] = (f"người dùng gửi 'DUYET {ma}' kèm ảnh email duyệt KSNB -> pull_keo_ve -> "
                           "pull_nap(env='test') -> người dùng gõ 'XAC NHAN <ma_xac_nhan>' -> pull_nap(env='prod')")
    return kq


# ───────────────────────── 0. LIỆT KÊ KỲ (chỉ đọc) ─────────────────────────
def liet_ke_ky(don_vi, thu_muc):
    """CHỈ ĐỌC available_metadata.json (không kéo, không ghi kế hoạch/nhật ký): các kỳ của một nguồn, để bot tự
    trả lời "tất cả các tháng / cả thư mục" rồi lập kế hoạch bằng `ten_file` — luật không kéo 'tất cả' của
    `lap_ke_hoach` và trần MAX_FILES giữ nguyên. Trả kỳ sắp tăng dần + file không suy được kỳ (sẽ bị bỏ)."""
    _pn, meta = _doc_meta()
    cd = [e for e in meta if e.get("fileName") and str(e.get("company") or "").upper() == str(don_vi or "").upper()
          and str(e.get("report_type") or "").lower() == str(thu_muc or "").lower()]
    if not cd:
        co = sorted({(e.get("company"), e.get("report_type")) for e in meta if e.get("company")})
        raise ValueError(f"không có nguồn {don_vi}/{thu_muc} trong danh sách nguồn (có: {co[:40]})")
    theo_ky, khong_ky = {}, []
    for e in cd:
        ky = _ky_cua_file(e)
        if ky:
            theo_ky.setdefault(ky, []).append(e["fileName"])
        else:
            khong_ky.append(e["fileName"])
    ds = [{"ky": k, "so_file": len(v), "file": sorted(v), "nam_cu": _ky_nam_cu(k)} for k, v in sorted(theo_ky.items())]
    tong = sum(x["so_file"] for x in ds)
    nay = sum(x["so_file"] for x in ds if not x["nam_cu"])
    return {"don_vi": cd[0]["company"], "thu_muc": cd[0]["report_type"], "so_ky": len(ds), "so_file": tong,
            "ky": ds, "khong_suy_duoc_ky": sorted(khong_ky), "tran_file_moi_ke_hoach": MAX_FILES,
            "so_ke_hoach_can": -(-tong // MAX_FILES) if tong else 0,
            "ky_nam_cu_cho_duyet": [x["ky"] for x in ds if x["nam_cu"]], "so_file_nam_nay": nay,
            "luu_y_nam_cu": "kỳ năm cũ chỉ NẠP được khi chính người yêu cầu gửi tin `DUYET <mã kế hoạch>` kèm ảnh "
                            "email duyệt của KSNB (chị Mai Thị Điệp)"}


# ───────────────────────── 2. KÉO VỀ ─────────────────────────
def keo_ve(plan_id, timeout=300):
    p = _doc_plan(plan_id)
    ten = {f["fileName"] for f in p["files"]}
    pn, _meta = _doc_meta()
    with _KhoaChay():
        kq = pn.keo(lambda e: e.get("company") == p["don_vi"] and e.get("report_type") == p["thu_muc"]
                    and e.get("fileName") in ten, log=lambda *_: None, refresh=False, timeout=timeout)
    ton_tai = [f["fileName"] for f in p["files"] if os.path.exists(_duong_dan(f))]
    p["trang_thai"]["keo"] = {"luc": _bay_gio().isoformat(timespec="seconds"), "xin": kq["xin"], "ve": kq["ve"],
                              "thieu": kq["thieu"]}
    _luu_plan(p)
    _audit("keo_ve", ma=p["ma_xac_nhan"], plan=plan_id, **{k: kq[k] if k != "thieu" else len(kq[k]) for k in ("xin", "ve", "thieu")})
    return {"plan_id": plan_id, "da_xin": kq["xin"], "da_ve": kq["ve"], "khong_ve_dung_han": kq["thieu"],
            "file_co_tren_dia": len(ton_tai),
            "luu_y": "file 'không về đúng hạn' vẫn dùng bản cũ trên đĩa nếu có; chưa từng có bản nào thì không nạp được"}


# ───────────────────────── 3. NẠP ─────────────────────────
def _luc(x):
    try:
        return datetime.datetime.fromisoformat(str(x.get("luc")))
    except (TypeError, ValueError):
        return None


def _xac_nhan_hop_le(p):
    """Đọc xác nhận do API ghi. Trả (ok, lý do).

    Phòng thủ 2 lớp (đợt 1, 15/10/2026): API đã chỉ ghi khi người gõ là chủ mã, tool vẫn so lại
    `xacnhan.user_id == ma_cua.user_id` — file xác nhận kiểu cũ (không user_id) bị từ chối. Đã dùng thì cờ nằm ở
    CẢ file xác nhận lẫn plan (`ma_da_dung_luc`): API có ghi đè file xác nhận khi người dùng gõ lại thì mã vẫn chết."""
    ma = p["ma_xac_nhan"]
    if p.get("ma_da_dung_luc"):
        return False, "mã xác nhận đã dùng một lần — lập kế hoạch mới"
    x = _doc_json(os.path.join(XN_DIR, ma + ".json"))
    if x is None:
        return False, f"chưa có xác nhận. Người dùng phải tự gõ: XAC NHAN {ma}"
    if x.get("da_dung"):
        return False, "mã xác nhận đã dùng một lần — lập kế hoạch mới"
    chu = _chu_ma(ma)
    if chu is None:
        return False, "mã chưa được hệ thống cấp cho người nào (thiếu ma_cua) — lập kế hoạch mới"
    if x.get("user_id") is None or x.get("user_id") != chu["user_id"]:
        return False, "xác nhận không phải của người đã yêu cầu kế hoạch này — chỉ người yêu cầu mới xác nhận được"
    luc = _luc(x)
    if luc is None:
        return False, "file xác nhận hỏng (thiếu thời điểm)"
    if luc < datetime.datetime.fromisoformat(p["tao_luc"]):
        return False, "xác nhận có trước khi lập kế hoạch"
    if _bay_gio() - luc > datetime.timedelta(minutes=XN_TTL_PHUT):
        return False, f"xác nhận đã quá {XN_TTL_PHUT} phút — người dùng gõ lại"
    return True, ""


def _anh_duyet_hop_le(p):
    """Kế hoạch có kỳ năm cũ (`can_anh_duyet`): cần `anh_duyet/<MÃ>.json` do API ghi khi chủ mã gửi `DUYET <MÃ>` kèm
    ảnh. Hợp lệ khi: cùng `ma`, `user_id` == chủ mã, có ≥1 tệp ảnh, `luc` sau `tao_luc` của plan. Trả (ok, lý do)."""
    ma = p["ma_xac_nhan"]
    x = _doc_json(os.path.join(ANH_DUYET_DIR, ma + ".json"))
    if x is None:
        return False, "chưa có ảnh email duyệt của KSNB cho kỳ năm cũ"
    chu = _chu_ma(ma)
    if x.get("ma") != ma:
        return False, "file ảnh duyệt không khớp mã kế hoạch"
    if chu is None or x.get("user_id") is None or x.get("user_id") != chu["user_id"]:
        return False, "ảnh duyệt không phải do người đã yêu cầu kế hoạch này gửi — chỉ người yêu cầu mới gửi được"
    tep = x.get("tep")
    if not isinstance(tep, list) or not any(isinstance(t, dict) and str(t.get("mime") or "").startswith("image/")
                                            for t in tep):
        return False, "tin DUYET không kèm tệp ảnh"
    luc = _luc(x)
    if luc is None or luc < datetime.datetime.fromisoformat(p["tao_luc"]):
        return False, "ảnh duyệt có trước khi lập kế hoạch — gửi lại cho kế hoạch này"
    return True, ""


def _danh_dau_da_dung(p):
    """Đánh dấu mã đã dùng TRƯỚC khi nạp PROD (trước 15/10 đánh dấu SAU → autofill ném lỗi/timeout giữa chừng thì
    mã còn sống, gọi lại được). Ghi plan trước (tool sở hữu), rồi file xác nhận (atomic). Ghi plan hỏng → ném lỗi,
    không nạp."""
    p["ma_da_dung_luc"] = _bay_gio().isoformat(timespec="seconds")
    _luu_plan(p)
    f = os.path.join(XN_DIR, p["ma_xac_nhan"] + ".json")
    x = _doc_json(f) or {}
    x["da_dung"] = True
    try:
        with open(f + ".tmp", "w", encoding="utf-8") as fh:
            json.dump(x, fh, ensure_ascii=False)
        os.replace(f + ".tmp", f)
    except OSError:
        pass                                            # cờ trong plan đã đủ chặn dùng lại


def nap(plan_id, env="test", timeout_moi_file=900):
    if env not in _PORT:
        raise ValueError("env phải là 'test' hoặc 'prod'")
    p = _doc_plan(plan_id)
    # plan do code trước 15/10 chiều lập (cờ LLM) mang khoá `anh_duyet_nam_cu` — vẫn phải qua cổng ảnh duyệt
    ky_cu = p.get("can_anh_duyet") or p.get("anh_duyet_nam_cu")
    if ky_cu:
        ok, ly_do = _anh_duyet_hop_le(p)
        if not ok:
            _audit("tu_choi_nam_cu", ma=p["ma_xac_nhan"], plan=plan_id, env=env, ky=ky_cu, ly_do=ly_do)
            return {"plan_id": plan_id, "env": env, "da_nap": False, "can_anh_duyet": ky_cu,
                    "ly_do": ly_do,
                    "huong_dan_cho_nguoi_dung": (f"Kỳ năm cũ cần ảnh email duyệt của KSNB: chính người yêu cầu gửi tin "
                                                 f"`DUYET {p['ma_xac_nhan']}` kèm ảnh")}
    hash_hien = {}
    thieu = []
    for f in p["files"]:
        d = _duong_dan(f)
        if os.path.exists(d):
            hash_hien[f["fileName"]] = _sha1(d)
        else:
            thieu.append(f["fileName"])
    if thieu:
        raise ValueError(f"{len(thieu)} file chưa có trên đĩa (chạy pull_keo_ve trước): {thieu[:5]}")
    if env == "prod":
        t = p["trang_thai"].get("test")
        if not t or not t.get("tat_ca_ok"):
            raise ValueError("chưa nạp TEST thành công cho kế hoạch này — nạp TEST trước")
        if t["hash"] != hash_hien:
            raise ValueError("file đã đổi kể từ lần nạp TEST — nạp lại TEST rồi xác nhận lại")
        ok, ly_do = _xac_nhan_hop_le(p)
        if not ok:
            return {"plan_id": plan_id, "env": "prod", "da_nap": False, "can_xac_nhan": True, "ly_do": ly_do,
                    "huong_dan_cho_nguoi_dung": f"Gõ đúng: XAC NHAN {p['ma_xac_nhan']}"}
    periods = {f["period"] for f in p["files"]}
    ket_qua, loi = [], []
    with _KhoaChay():
        if env == "prod":                               # dùng một lần — đánh dấu TRƯỚC khi đụng DB
            p = _doc_plan(plan_id)                      # đọc lại trong khoá: lượt song song vừa dùng mã thì thấy ngay
            ok, ly_do = _xac_nhan_hop_le(p)
            if not ok:
                return {"plan_id": plan_id, "env": "prod", "da_nap": False, "can_xac_nhan": True, "ly_do": ly_do,
                        "huong_dan_cho_nguoi_dung": f"Gõ đúng: XAC NHAN {p['ma_xac_nhan']}"}
            _danh_dau_da_dung(p)
            _audit("bat_dau_nap_prod", ma=p["ma_xac_nhan"], plan=plan_id)
        truoc = _anh_chup(env, p["don_vi"], periods)
        envp = {**os.environ, "DATABASE_URL":
                f"postgresql://tc:tc_%24production@localhost:{_PORT[env]}/tc_dashboard"}
        for f in p["files"]:
            cmd = [AGENT_PY, "scripts/agent_cli.py", "autofill", _duong_dan(f), "--period", f["period"]]
            try:
                r = subprocess.run(cmd, cwd=AGENT_DIR, env=envp, capture_output=True, text=True,
                                   timeout=timeout_moi_file)
                dong = ((r.stdout or r.stderr or "").strip().splitlines() or [""])[-1]
                try:
                    js = json.loads(dong)
                    ok = bool(js.get("ok"))
                    ly = None if ok else str(js.get("error") or js.get("skip") or "ok:false")[:200]
                except ValueError:
                    ok, ly = False, f"rc={r.returncode}: {dong[:160]}"
            except subprocess.TimeoutExpired:
                ok, ly = False, f"quá {timeout_moi_file}s"
            ket_qua.append({"file": f["fileName"], "ok": ok, "ly_do_loi": ly})
            if not ok:
                loi.append(f["fileName"])
        sau = _anh_chup(env, p["don_vi"], periods)
    diff = _so_sanh(truoc, sau)
    p["trang_thai"][env] = {"luc": _bay_gio().isoformat(timespec="seconds"), "tat_ca_ok": not loi,
                            "hash": hash_hien, "so_file": len(ket_qua), "so_loi": len(loi),
                            "so_nhom_doi": diff["so_nhom_doi"]}
    _luu_plan(p)
    _audit("nap", ma=p["ma_xac_nhan"], plan=plan_id, env=env, so_file=len(ket_qua), so_loi=len(loi), so_nhom_doi=diff["so_nhom_doi"])
    return {"plan_id": plan_id, "env": env, "da_nap": True, "so_file": len(ket_qua), "so_file_loi": len(loi),
            "ket_qua_tung_file": ket_qua, "so_sanh_truoc_sau": diff,
            "buoc_tiep": (f"Nếu số liệu ổn: người dùng gõ 'XAC NHAN {p['ma_xac_nhan']}' rồi gọi pull_nap(env='prod')"
                          if env == "test" and not loi else None)}


def trang_thai(plan_id):
    p = _doc_plan(plan_id)
    return {"plan_id": plan_id, "don_vi": p["don_vi"], "thu_muc": p["thu_muc"], "so_file": len(p["files"]),
            "tao_luc": p["tao_luc"], "trang_thai": {k: (None if v is None else
                                                          {kk: vv for kk, vv in v.items() if kk != "hash"})
                                                    for k, v in p["trang_thai"].items()}}

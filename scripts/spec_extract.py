# -*- coding: utf-8 -*-
"""ENGINE TRÍCH XUẤT THEO SPEC — thay cho việc viết 1 file Python cho mỗi nguồn báo cáo.

VÌ SAO CÓ FILE NÀY (2026-08-08): mọi deriver trước đây (`derive_hqkd_ngay.py`,
`derive_congno_tuoino.py`, `derive_vhkd_kqkd.py`…) đọc cột theo CHỈ SỐ CỨNG. Kế toán chèn/xoá 1 cột
là sai số im lặng, và mỗi lần mapping đổi là một vòng sửa Python → test → deploy. Đợt VHKD/XDV có
~10 nguồn nên chi phí đó nhân lên. Engine này đọc một **spec JSON** mô tả nguồn và ánh xạ cột, nên
đổi mapping = sửa JSON, không đụng code.

NGUYÊN TẮC SỐ 1 — DÒ CỘT THEO **TÊN HEADER**, KHÔNG THEO CHỮ CỘT.
Riêng điều này diệt cả một lớp lỗi đã gặp thật: file mapping ghi kho xe B2B "Trạng thái >30/45
ngày" ở cột AK, file thật để ở AM (AK là "Số ngày tồn kho thực tế"); kho xe B2C ghi "Tên KH" cột O,
thật ra ở P. Có thể khai `"cot_du_phong": "AM"` làm phao khi header rỗng, nhưng tên luôn thắng.

SPEC ĐẶT Ở: `Dashboard_Agent/extract_specs/<id>.json` (KHÔNG để trong `memory/` — thư mục đó
bị .gitignore vì là state runtime của agent; spec là CẤU HÌNH NGUỒN, phải được version). Đây là file CẤU HÌNH TIN CẬY của nội
bộ (không nhận từ ngoài) — `dan_xuat` dùng eval trong sandbox hẹp; đừng mở cho input người dùng.

CẤU TRÚC SPEC (khoá tiếng Việt cho kế toán/BA đọc được):
{
  "id": "vhkd_kqkd",                     // trùng tên file
  "report_type": "KDVH",                 // ghi vào raw_rows.report_type
  "row_index_base": 6400000,             // dải row_index riêng, tránh đụng deriver khác
  "khoi": "Khối KD Vinfast - Showroom",  // hằng số (hoặc bỏ, nếu suy từ cost_center)
  "nguon": {
    "folder": "SRVF/baocaokqkd",         // dưới received_reports/
    "file_glob": "*.xlsx",
    "sheet": {"batdau": "CHI TIẾT XHĐ"}  // | {"ten": "..."} | {"chua": "..."} | {"so": 0}
  },
  "header": {"dong": 1},                 // | {"dong": [1,2]} gộp 2 dòng header (lấy ô đầu khác rỗng)
  "dong_bat_dau": 3,                     // tuỳ chọn — mặc định = dòng header cuối + 1
  "chieu_tu_ten_file": {"dim2": {"regex": "Xuathoadon_(B2B|B2C|GF)", "hoa": true}},
  "ngay_tu_ten_file": {"regex": "M\\.(\\d{4})\\.(\\d{1,2})\\.(\\d{1,2})", "thu_tu": "ymd"},
  "cot": {                               // đích -> cách lấy. Đích: ngay/cost_center/cong_ty/
    "cost_center": {"header": "Tên DVCS", "chuan_hoa": "sr_showroom"},   // amount/amount2/dim1..3/
    "ngay":   {"header": "Ngày hóa đơn", "kieu": "date"},                // payload.<khoá bất kỳ>
    "amount": {"header": "Giá bán", "kieu": "so", "he_so": 1e-9}
  },
  "ban_ghi": "moi_dong",                 // | "moi_cot_gia_tri" (xem dưới)
  "cot_gia_tri": [                       // chỉ dùng khi ban_ghi = "moi_cot_gia_tri":
    {"header": "Công nợ trong hạn", "dim1": "Trong hạn", "he_so": 1e-9}
  ],                                     // -> mỗi dòng nguồn đẻ N bản ghi, amount lấy từng cột
  "loc": [{"cot": "ngay", "dieu_kien": "khac_rong"}],
  "dan_xuat": {"payload.lng": "amount - payload.gia_von"},
  "payload_them": {"unit": "ty"}
}

`ban_ghi`:
  - "moi_dong"        : 1 dòng nguồn -> 1 bản ghi (bảng chi tiết: xuất hoá đơn, claim, kho xe).
  - "moi_cot_gia_tri" : 1 dòng nguồn -> N bản ghi, mỗi cột giá trị thành 1 dòng có dim1 riêng
                        (bảng ma trận: "Tuổi nợ phải thu" có sẵn cột trong hạn/1-30/>30-90/…).

Ghi DB idempotent theo `source_file` = "<FOLDER>::<tên file>" — cùng quy ước mọi deriver khác.
`--write` mới ghi; mặc định dry-run in ra tổng hợp để đối chiếu với file gốc.

HOOK (khi spec không tả nổi): thêm hàm vào `_CHUAN_HOA` bên dưới rồi gọi bằng tên trong
`"chuan_hoa"`. Hook trả str, hoặc dict để set nhiều trường cùng lúc (vd cost_center + cong_ty).
"""
import argparse
import calendar
import datetime as dt
import glob
import json
import os
import re
import sys
import unicodedata

import openpyxl
import psycopg
from openpyxl.utils import column_index_from_string

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(_ROOT, ".env"))

DB_URL = os.environ.get("DATABASE_URL")
SPEC_DIR = os.path.join(_ROOT, "extract_specs")
REPORTS_DIR = os.path.normpath(os.path.join(_ROOT, "..", "Connect_VPS", "received_reports"))

_CHUAN_TRUONG = {"ngay", "cong_ty", "khoi", "cost_center", "amount", "amount2",
                 "dim1", "dim2", "dim3"}


# ─────────────────────────── tiện ích chuẩn hoá ───────────────────────────
def _nd(s):
    """Bỏ dấu + bỏ ký tự không phải chữ/số + thường hoá — dùng để so tên header/tên đơn vị."""
    s = str(s if s is not None else "").strip().lower().replace("đ", "d")
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s)


def _so(v, he_so=1.0):
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) * he_so
    s = str(v or "").strip().replace(" ", "")
    if not s:
        return 0.0
    # '1.234.567,89' (VN) và '1,234,567.89' (EN) đều gặp trong file kế toán
    s = s.replace(".", "").replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        return float(s) * he_so
    except ValueError:
        return 0.0


def _date(v):
    """-> 'YYYY-MM-DD' hoặc None. File có cả datetime, 'dd/mm/yyyy' lẫn 'yyyy-mm-dd'."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
    else:
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if not m:
            return None
        y, mo, d = (int(x) for x in m.groups())
    try:
        return dt.date(y, mo, d).isoformat()
    except ValueError:
        return None


# ─────────────────────────── hook chuẩn hoá ───────────────────────────
_CC_CACHE = {}


def _cc_showroom(ten):
    """Tên đơn vị trong file -> {cost_center, cong_ty} của khối Vinfast - Showroom.

    File ghi "Showroom OceanPark"/"Showroom Uông Bí", master ghi "Vinfast Ocean Park" -> bỏ tiền
    tố showroom/vinfast ở CẢ hai phía rồi so. Bản "(61)" (pháp nhân Xanh Vĩnh Phúc) chuẩn hoá
    thành "...61" nên không đụng bản gốc. "Vinfast B2B" là đội bán B2B tập trung -> B2B_SR.
    """
    return _cc_theo_khoi(ten, "Khối KD Vinfast - Showroom", {"b2b": ("B2B_SR", "TC")})


def _cc_xdv(ten):
    return _cc_theo_khoi(ten, "Khối KD Vinfast - XDV", {})


def _cc_theo_khoi(ten, khoi, alias):
    if khoi not in _CC_CACHE:
        sys.path.insert(0, os.path.join(_ROOT, "..", "AI_coding", "tc-admin-api"))
        from app.master_data import loader as master
        m = {}
        for cc in master.master_data().get("costCenters", []):
            if (cc.get("khoi") or "") != khoi:
                continue
            m[_bo_tien_to(_nd(cc.get("ten")))] = (str(cc.get("ma") or "").strip(),
                                                 master.resolve_company_code(cc.get("congTy") or ""))
        _CC_CACHE[khoi] = m
    key = _bo_tien_to(_nd(ten))
    ma, cty = _CC_CACHE[khoi].get(key) or alias.get(key) or (None, None)
    return {"cost_center": ma, "cong_ty": cty} if ma else {"_khong_map": str(ten or "").strip()}


def _bo_tien_to(n):
    for p in ("showroom", "vinfast", "xuongdichvu", "xdv"):
        if n.startswith(p):
            n = n[len(p):]
    return n


_CHUAN_HOA = {
    "sr_showroom": _cc_showroom,
    "xdv": _cc_xdv,
    "hoa": lambda v: str(v or "").strip().upper() or None,
    "cat": lambda v: str(v or "").strip() or None,
}


# ─────────────────────────── đọc spec & sheet ───────────────────────────
def load_spec(ref):
    path = ref if os.path.isfile(ref) else os.path.join(SPEC_DIR, f"{ref}.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _chon_sheet(wb, cfg):
    cfg = cfg or {}
    if "so" in cfg:
        return wb.sheetnames[int(cfg["so"])]
    for name in wb.sheetnames:
        n = _nd(name)
        if "ten" in cfg and n == _nd(cfg["ten"]):
            return name
        if "batdau" in cfg and n.startswith(_nd(cfg["batdau"])):
            return name
        if "chua" in cfg and _nd(cfg["chua"]) in n:
            return name
    return None


def _map_header(rows, dong):
    """{tên header đã chuẩn hoá: chỉ số cột 0-based}. `dong` là số hoặc list số (1-based).

    Gộp nhiều dòng: lấy ô KHÁC RỖNG ĐẦU TIÊN theo thứ tự khai — file Claim để nhóm cột ở dòng 1
    ("XE"/"PIN") còn tên cột thật ở dòng 2 (Z "SR", AA "Trạng thái", AB "Số tiền").
    """
    dongs = dong if isinstance(dong, list) else [dong]
    ncol = max((len(rows[d - 1]) for d in dongs if d - 1 < len(rows)), default=0)
    out, theo_dong = {}, {d: {} for d in dongs}
    for j in range(ncol):
        xong = False
        for d in dongs:
            r = rows[d - 1] if d - 1 < len(rows) else ()
            v = r[j] if j < len(r) else None
            n = _nd(v)
            if n:
                theo_dong[d].setdefault(n, j)
                if not xong:
                    out.setdefault(n, j)
                    xong = True
    return out, theo_dong


def _resolve_cot(spec, hmap, warn):
    """{đích: (chỉ số cột, cfg)} — dò theo TÊN, phao là `cot_du_phong` (chữ cột)."""
    out = {}
    for dich, cfg in (spec.get("cot") or {}).items():
        j = _tim_cot(hmap, cfg, dich, warn)
        if j is None:
            continue
        out[dich] = (j, cfg)
    return out


def _tim_cot(hmaps, cfg, nhan, warn):
    """Chỉ số cột cho 1 khai báo. `header` nhận CHUỖI hoặc LIST tên ứng viên (thử theo thứ tự).

    Vì sao cần list: cùng một cột nhưng file đổi tên giữa các kỳ — báo cáo Claim ghi "Check" ở
    T1/T2 rồi đổi thành "Trạng thái" từ T3. Khai `["Check", "Trạng thái"]` chạy được cả 6 file.

    Vì sao cần `dong_header`: T1/T2 còn có một cột MỒI ở dòng 1 cũng tên "Trạng thái" nhưng chứa
    ĐỊA CHỈ khách. Không giới hạn dòng header thì trạng thái claim biến thành "Thôn Xám, Xã Vĩnh
    Phú…" — sai mà vẫn chạy trơn. Khai `"dong_header": 2` để chỉ dò trên đúng dòng tiêu đề thật.
    """
    hmap, theo_dong = hmaps
    if cfg.get("dong_header"):
        hmap = theo_dong.get(cfg["dong_header"], {})
    ten = cfg.get("header")
    ung_vien = ten if isinstance(ten, list) else ([ten] if ten else [])
    for t in ung_vien:
        j = hmap.get(_nd(t))
        if j is not None:
            return j
    if cfg.get("cot") or cfg.get("cot_du_phong"):
        chu = cfg.get("cot") or cfg["cot_du_phong"]
        if cfg.get("cot_du_phong") and ung_vien:
            warn.append(f"cột '{ung_vien[0]}' không thấy theo tên -> dùng phao {chu}")
        return column_index_from_string(chu) - 1
    if cfg.get("bat_buoc", True):
        warn.append(f"THIẾU cột bắt buộc '{ten}' ({nhan})")
    return None


def _lay_o(row, j, cfg):
    v = row[j] if j < len(row) else None
    kieu = cfg.get("kieu", "text")
    if kieu == "so":
        return _so(v, float(cfg.get("he_so", 1.0)))
    if kieu == "date":
        return _date(v)
    s = str(v).strip() if v is not None else None
    return s or None


def _dat(rec, dich, val):
    if dich.startswith("payload."):
        rec.setdefault("payload", {})[dich[len("payload."):]] = val
    else:
        rec[dich] = val


def _lay(rec, dich):
    if dich.startswith("payload."):
        return (rec.get("payload") or {}).get(dich[len("payload."):])
    return rec.get(dich)


def _qua_loc(rec, loc):
    for f in loc or []:
        v = _lay(rec, f["cot"])
        dk, gt = f.get("dieu_kien", "khac_rong"), f.get("gia_tri")
        if dk == "khac_rong" and (v is None or v == "" or v == 0):
            return False
        if dk == "bang" and str(v) != str(gt):
            return False
        if dk == "khac" and str(v) == str(gt):
            return False
        if dk == "thuoc" and str(v) not in [str(x) for x in gt]:
            return False
        if dk == "khong_thuoc" and str(v) in [str(x) for x in gt]:
            return False
        if dk == "lon_hon" and not (isinstance(v, (int, float)) and v > gt):
            return False
        if dk == "nho_hon" and not (isinstance(v, (int, float)) and v < gt):
            return False
        if dk == "chua" and _nd(gt) not in _nd(v):
            return False
    return True


def _dan_xuat(rec, cong_thuc):
    """eval trong sandbox hẹp — spec là cấu hình nội bộ tin cậy (xem docstring đầu file)."""
    if not cong_thuc:
        return
    env = {k: v for k, v in rec.items() if isinstance(v, (int, float))}
    env.update({f"payload.{k}": v for k, v in (rec.get("payload") or {}).items()})
    safe = {re.sub(r"\W", "_", k): (v or 0) for k, v in env.items()}
    for dich, bt in cong_thuc.items():
        try:
            _dat(rec, dich, eval(re.sub(r"\W", "_", bt) if "." in bt else bt,  # noqa: S307
                                {"__builtins__": {}}, safe))
        except Exception as ex:                                    # noqa: BLE001
            _dat(rec, dich, None)
            rec.setdefault("_loi", []).append(f"{dich}: {ex}")


def ngay_tu_ten_file(spec, path):
    """-> ('YYYY-MM-DD' | None, [cảnh báo]). Dùng cho bảng ẢNH CHỤP (không có cột ngày từng dòng).

    `ky_tu_ten_file` tách RIÊNG regex năm và tháng — file công nợ đặt tên
    "…M.2026.07.22_Baocaocongnophaithu_T1.xlsx": cụm 2026.07.22 là NGÀY LẬP báo cáo (giống hệt
    nhau ở mọi kỳ), kỳ thật nằm ở hậu tố _T1/_T7. Lấy nhầm cụm đầu là dồn cả 12 tháng vào tháng 7.
    `ngay_tu_ten_file` dùng khi tên file có ngày chốt thật (kho xe B2B: "…2026.7.31.KHO XE.xlsx").
    """
    ten = os.path.basename(path)
    warn = []
    if spec.get("ky_tu_ten_file"):
        c = spec["ky_tu_ten_file"]
        mn, mt = re.search(c["regex_nam"], ten), re.search(c["regex_thang"], ten)
        if mn and mt:
            y, mo = int(mn.group(1)), int(mt.group(1))
            return dt.date(y, mo, calendar.monthrange(y, mo)[1]).isoformat(), warn
        warn.append(f"không dò được kỳ (năm/tháng) từ tên file: {ten}")
    if spec.get("ngay_tu_ten_file"):
        m = re.search(spec["ngay_tu_ten_file"]["regex"], ten)
        if m:
            try:
                return dt.date(*(int(m.group(i + 1)) for i in range(3))).isoformat(), warn
            except ValueError:
                warn.append(f"ngày trong tên file không hợp lệ: {m.group(0)}")
        else:
            warn.append(f"không dò được ngày từ tên file: {ten}")
    return None, warn


def loc_file_moi_nhat(spec, files):
    """Mỗi KỲ chỉ giữ file có ngày chốt MỚI NHẤT.

    Kho xe B2B có 2 ảnh chụp cùng tháng 7 ("…7.24.KHO XE" và "…7.31.KHO XE"). Nạp cả hai là
    ĐẾM ĐÔI tồn kho tháng 7 (mỗi VIN xuất hiện 2 lần) — số tồn phình gần gấp đôi mà không có
    cảnh báo nào. Bật `"moi_ky_lay_file_moi_nhat": true` trong spec cho mọi nguồn ảnh chụp.
    """
    if not spec.get("moi_ky_lay_file_moi_nhat"):
        return files, []
    giu, bo = {}, []
    for f in files:
        ngay, _ = ngay_tu_ten_file(spec, f)
        if not ngay:
            giu[f] = f            # không suy được kỳ -> giữ nguyên, đừng im lặng loại
            continue
        ky = ngay[:7]
        cu = giu.get(ky)
        if cu is None or ngay > ngay_tu_ten_file(spec, cu)[0]:
            if cu is not None:
                bo.append(os.path.basename(cu))
            giu[ky] = f
        else:
            bo.append(os.path.basename(f))
    return sorted(giu.values()), bo


# ─────────────────────────── trích 1 file ───────────────────────────
def extract_file(spec, path):
    warn, recs = [], []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = _chon_sheet(wb, (spec.get("nguon") or {}).get("sheet"))
        if not sheet:
            return [], [f"không chọn được sheet (có: {wb.sheetnames})"]
        ws = wb[sheet]
        hdr_cfg = spec.get("header") or {"dong": 1}
        if hdr_cfg.get("khong_co"):
            # Sheet KHÔNG có dòng tiêu đề (vd "Tồn kho xe vật lý": dữ liệu chạy thẳng từ dòng 2,
            # 3 cột không tên). Khi đó mọi khai báo cột PHẢI dùng `"cot": "<chữ cột>"`.
            max_hdr, hmap = 0, ({}, {})
        else:
            hdr_dong = hdr_cfg.get("dong", 1)
            max_hdr = max(hdr_dong) if isinstance(hdr_dong, list) else hdr_dong
            head_rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=max_hdr,
                                                       values_only=True)]
            hmap = _map_header(head_rows, hdr_dong)
        cot = _resolve_cot(spec, hmap, warn)

        chieu = {}
        for dich, cfg in (spec.get("chieu_tu_ten_file") or {}).items():
            m = re.search(cfg["regex"], os.path.basename(path), re.I)
            v = (m.group(int(cfg.get("nhom", 1))) if m else cfg.get("mac_dinh"))
            chieu[dich] = (str(v).upper() if v and cfg.get("hoa") else v)

        # CHỐT CHẶN LAYOUT: nguồn chỉ đúng khi các ô mốc khớp. Bắt buộc với spec dùng CHỮ CỘT —
        # file công nợ T1 và T7 tuy cùng tên sheet nhưng bố cục KHÁC HẲN (T1 không có phân tách
        # kênh, header ở dòng 5). Không có chốt này thì T1 vẫn "chạy được" và đẻ ra 9,86 tỷ vô
        # nghĩa, im lặng — kiểu sai nguy hiểm nhất.
        for ktr in spec.get("kiem_tra_o") or []:
            thuc = ws[ktr["o"]].value
            if _nd(ktr.get("bang")) not in _nd(thuc):
                return [], [f"BỎ QUA — layout khác spec: ô {ktr['o']} = "
                            f"{str(thuc)[:40]!r}, cần chứa {ktr.get('bang')!r}"]

        ngay_file, w2 = ngay_tu_ten_file(spec, path)
        warn.extend(w2)

        bat_dau = spec.get("dong_bat_dau") or (max_hdr + 1)
        gia_tri_cols = spec.get("cot_gia_tri") or []
        gt_idx = [(_tim_cot(hmap, c, f"cột giá trị {c.get('dim1') or c.get('header')}", warn), c)
                  for c in gia_tri_cols]
        gt_idx = [(j, c) for j, c in gt_idx if j is not None]

        khong_map, bo_loc = {}, 0
        for row in ws.iter_rows(min_row=bat_dau, values_only=True):
            base = {"payload": dict(spec.get("payload_them") or {})}
            if spec.get("khoi"):
                base["khoi"] = spec["khoi"]
            base.update(spec.get("chieu_co_dinh") or {})   # vd {"dim2": "B2B"} cho file 1 kênh
            base.update({k: v for k, v in chieu.items()})
            if ngay_file:
                base["ngay"] = ngay_file
            for dich, (j, cfg) in cot.items():
                val = _lay_o(row, j, cfg)
                hook = cfg.get("chuan_hoa")
                if hook and val is not None:
                    res = _CHUAN_HOA[hook](val)
                    if isinstance(res, dict):
                        # Hook hỏng KHÔNG loại dòng ngay: phải để `loc` chạy trước, nếu không mọi
                        # dòng RÁC (vùng dán thừa, không có ngày) sẽ bị gộp vào cảnh báo "không map
                        # được đơn vị" và che mất đơn vị hỏng THẬT. Đánh dấu rồi xử lý sau bộ lọc.
                        if res.get("_khong_map"):
                            # `giu_khi_khong_map`: dòng KHÔNG khớp danh mục vẫn là dòng THẬT và
                            # phải giữ để tổng không hụt — vd 16 xe tồn đứng tên "CHI NHÁNH
                            # VINFAST HÀ NỘI" trong báo cáo tồn vật lý: không thuộc SR nào nhưng
                            # là xe có thật. Giữ lại, cost_center để trống, tên thô nằm ở payload.
                            if not cfg.get("giu_khi_khong_map"):
                                base["_khong_map"] = res["_khong_map"]
                        for k2, v2 in res.items():
                            if k2 != "_khong_map":
                                base[k2] = v2
                        continue
                    val = res
                _dat(base, dich, val)
            if all(base.get(k) in (None, "") for k in _CHUAN_TRUONG if k != "khoi") \
                    and not (base.get("payload") or {}):
                continue
            # DỪNG QUÉT khi gặp dòng thoả `dung_khi` — sheet "CN theo đơn vị" có dòng "Total" rồi
            # BÊN DƯỚI còn một bảng phụ (công nợ quá hạn đã có COC) cũng liệt kê tên Showroom ở
            # cột A. Không dừng là gộp luôn bảng phụ vào -> cộng đôi mà không có dấu hiệu gì.
            if spec.get("dung_khi") and _qua_loc(base, spec["dung_khi"]):
                break
            outs = []
            if spec.get("ban_ghi") == "moi_cot_gia_tri":
                for j, c in gt_idx:
                    r2 = json.loads(json.dumps(base))
                    r2["amount"] = _so(row[j] if j < len(row) else None,
                                       float(c.get("he_so", 1.0)))
                    for k2 in ("dim1", "dim2", "dim3"):
                        if c.get(k2):
                            r2[k2] = c[k2]
                    outs.append(r2)
            else:
                outs.append(base)
            for r2 in outs:
                _dan_xuat(r2, spec.get("dan_xuat"))
                hong = r2.pop("_khong_map", None)
                if not _qua_loc(r2, spec.get("loc")):
                    bo_loc += 1
                elif hong:
                    khong_map[hong] = khong_map.get(hong, 0) + 1
                else:
                    recs.append(r2)
        if bo_loc:
            warn.append(f"bỏ {bo_loc} dòng không qua bộ lọc")
        if khong_map:
            warn.append("KHÔNG MAP ĐƯỢC đơn vị (đã qua bộ lọc, tức là dòng THẬT): "
                        + ", ".join(f"{k} ({v} dòng)"
                                    for k, v in sorted(khong_map.items(), key=lambda x: -x[1])[:10]))
        return recs, warn
    finally:
        wb.close()


def _source_id(path):
    parts = os.path.normpath(path).split(os.sep)
    return f"{parts[-3]}::{os.path.basename(path)}" if len(parts) >= 3 else os.path.basename(path)


def _ghi(spec, path, recs):
    source_file = _source_id(path)
    conn = psycopg.connect(DB_URL)
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM raw_rows WHERE report_type=%s AND source_file=%s",
                    (spec["report_type"], source_file))
        by_ky, thieu_ky = {}, 0
        for r in recs:
            if not r.get("ngay"):
                thieu_ky += 1
                continue
            by_ky.setdefault(r["ngay"][:7], []).append(r)
        rows, skipped, i = [], [], 0
        base_i = int(spec.get("row_index_base", 6500000))
        for period, items in sorted(by_ky.items()):
            cur.execute("SELECT id FROM datasets WHERE kind='month' AND period=%s "
                        "ORDER BY created_at DESC LIMIT 1", (period,))
            got = cur.fetchone()
            if not got:
                skipped.append(period)
                continue
            for r in items:
                i += 1
                rows.append((got[0], spec["report_type"], base_i + i, r.get("ngay"),
                             r.get("cong_ty"), r.get("khoi"), r.get("cost_center"), period,
                             r.get("amount"), r.get("amount2"), r.get("dim1"), r.get("dim2"),
                             r.get("dim3"),
                             json.dumps(r.get("payload") or {}, ensure_ascii=False), source_file))
        if rows:
            cur.executemany(
                "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
                "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, "
                "source_file) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", rows)
        conn.commit()
        return {"written": len(rows),
                **({"bo_qua_chua_co_dataset": skipped} if skipped else {}),
                **({"bo_qua_khong_co_ngay": thieu_ky} if thieu_ky else {})}
    finally:
        conn.close()


def run(spec, path, write=False):
    recs, warn = extract_file(spec, path)
    by_ky = {}
    for r in recs:
        by_ky.setdefault((r.get("ngay") or "?")[:7], []).append(r)
    out = {"file": os.path.basename(path), "dong": len(recs),
           "ky": {k: {"dong": len(v),
                      "amount": round(sum(x.get("amount") or 0 for x in v), 6)}
                  for k, v in sorted(by_ky.items())}}
    if warn:
        out["canh_bao"] = warn
    if write and recs:
        out.update(_ghi(spec, path, recs))
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Trích xuất báo cáo theo spec JSON.")
    ap.add_argument("spec", help="id spec (memory/extract_specs/<id>.json) hoặc đường dẫn file")
    ap.add_argument("--file", help="chỉ chạy 1 file (mặc định: quét cả folder trong spec)")
    ap.add_argument("--write", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    sp = load_spec(a.spec)
    if a.file:
        files, bo = [a.file], []
    else:
        ng = sp["nguon"]
        files = sorted(glob.glob(os.path.join(REPORTS_DIR, ng["folder"],
                                              ng.get("file_glob", "*.xlsx"))))
        files, bo = loc_file_moi_nhat(sp, files)
    ket_qua = [run(sp, f, write=a.write) for f in files]
    if bo:
        ket_qua.append({"_bo_qua_anh_chup_cu": bo})
    print(json.dumps(ket_qua, ensure_ascii=False, indent=2))

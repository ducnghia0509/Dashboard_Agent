# -*- coding: utf-8 -*-
"""TÁCH KHỐI LOG LỖI DÁN TỪ EXCEL thành `nguon_log` của phiếu (schema v1) — TẤT ĐỊNH, không LLM (15/10/2026).

VÌ SAO có module riêng: Agent tổng (`quan_ly`) nhận nguyên khối người dùng dán từ file log lỗi dashboard
(Excel -> clipboard = TSV, ô nhiều dòng nằm trong ngoặc kép). Để LLM tự tách thì nó đọc nhầm cột, làm tròn số,
bỏ dòng; ở đây đọc bằng `csv` dialect `excel-tab` (đúng cách Excel chép ra) rồi chuẩn hoá từng ô.

Bố cục cột log (mẫu Xe tải 15/10/2026, `scripts/du_lieu_test/log_mau_xetai_so_gia.tsv`):
  Tên User · STT · Cost Center · Dashboard mẹ · Chỉ tiêu con · Chỉ tiêu · Kỳ, tháng · Số báo cáo · Số dashboard ·
  Chênh lệch (I-J) · Mô tả lỗi · Phân tích nghiệp vụ xác nhận loại · IT fix lỗi · Trạng thái · Ghi chú
Có dòng tiêu đề thì đặt cột theo TÊN (thứ tự cột đổi vẫn đúng); không có thì theo thứ tự trên.

Quy ước (giám sát chốt 15/10/2026):
  - Ô trống ở Tên User / Dashboard mẹ KẾ THỪA dòng trên (người báo chỉ điền dòng đầu nhóm).
  - Số: " 12,345,678,901 " -> 12345678901; " (7)" -> -7; " -   " -> gia_tri null (ô '-' của định dạng kế toán =
    dashboard trống); "1.7 triệu USD" -> 1700000 usd. Không làm tròn (Decimal).
  - ĐƠN VỊ KHÔNG ĐOÁN theo ngữ cảnh: chỉ lấy từ chữ trong ô (USD/đ/%/xe) hoặc từ tên chỉ tiêu rõ nghĩa
    (doanh thu/chi phí… -> đồng; tỷ lệ -> pct; "xe" trong chỉ tiêu số lượng -> xe). "Giá trị hàng đã nhập kho
    2,500,000" (thật ra USD) -> don_vi 'khac' + ghi_chu, để bot tìm lỗi xác định từ mapping.
  - TRÙNG chỉ khi cùng màn + chỉ tiêu + kỳ. "Giống lỗi N" chỉ là LIÊN KẾT (`lien_quan`), không phải trùng —
    log mẫu có 2 dòng T7 "Giống lỗi 12/13" trỏ về dòng T8 cùng chỉ tiêu: kỳ khác = lỗi khác cần đo riêng.
  - `dong_mapping` = "dòng 10", "dòng 22-29" trong mô tả/phân tích -> "<sheet mới nhất của tài liệu mapping
    của màn>!<dòng>" (tra `mapping_index.json`); không tra được sheet thì "?!<dòng>".
"""
import csv
import io
import json
import os
import re
import unicodedata
from decimal import Decimal, InvalidOperation

_ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
MAPPING_INDEX = os.path.join(_ROOT, "mapping_index.json")
MAX_DONG = 60

COT = ["nguoi_bao", "stt", "cost_center", "dashboard_me", "chi_tieu_con", "chi_tieu", "ky", "so_bao_cao",
       "so_dashboard", "chenh_lech", "mo_ta", "phan_tich", "it_fix", "trang_thai_log", "ghi_chu"]
# tiền tố tên cột (đã bỏ dấu, chữ thường) -> khoá; xét theo thứ tự (chỉ tiêu con TRƯỚC chỉ tiêu).
_TEN_COT = [("ten user", "nguoi_bao"), ("nguoi bao", "nguoi_bao"), ("stt", "stt"), ("cost center", "cost_center"),
            ("dashboard me", "dashboard_me"), ("chi tieu con", "chi_tieu_con"), ("chi tieu", "chi_tieu"),
            ("ky", "ky"), ("so bao cao", "so_bao_cao"), ("so dashboard", "so_dashboard"),
            ("chenh lech", "chenh_lech"), ("mo ta", "mo_ta"), ("phan tich", "phan_tich"), ("it fix", "it_fix"),
            ("trang thai", "trang_thai_log"), ("ghi chu", "ghi_chu")]
_KE_THUA = ("nguoi_bao", "dashboard_me")
_TU_TIEN = re.compile(r"doanh thu|chi phi|loi nhuan|gia von|tien|cong no|phai thu|phai tra|thue|lai vay|ngan sach")


def kd(s):
    """Bỏ dấu + chữ thường + gộp khoảng trắng (đ -> d)."""
    s = unicodedata.normalize("NFD", str(s or "")).replace("đ", "d").replace("Đ", "D")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"\s+", " ", s).strip()


# ───────────────────────── từng ô ─────────────────────────
def doc_ky(s):
    """'T8/2026' -> '2026-08'; 'Ngày 06/08/2026' -> '2026-08-06'. Không nhận ra -> None (giữ nguyên văn ở nơi gọi)."""
    t = kd(s)
    if not t:
        return None
    m = re.fullmatch(r"(?:ngay\s*)?(\d{1,2})[/.\-](\d{1,2})[/.\-](20\d{2})", t)
    if m and 1 <= int(m[2]) <= 12 and 1 <= int(m[1]) <= 31:
        return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}"
    m = re.fullmatch(r"(20\d{2})-(\d{2})(?:-(\d{2}))?", t)
    if m:
        return t
    m = re.fullmatch(r"(?:t|thang|th)?\s*(\d{1,2})\s*[/.\-]\s*(20\d{2})", t)
    if m and 1 <= int(m[1]) <= 12:
        return f"{m[2]}-{int(m[1]):02d}"
    return None


def _so_tu_chuoi(t, co_he_so):
    """Chuỗi số (đã bỏ chữ) -> Decimal. Dấu nghìn/thập phân: nhiều dấu cùng loại hoặc nhóm đúng 3 chữ số = nghìn;
    một dấu có hệ số (triệu/tỷ) hoặc không đúng 3 chữ số sau = thập phân. None nếu không đọc được."""
    t = t.replace(" ", "")
    if not re.fullmatch(r"\d[\d.,]*", t):
        return None
    cham, phay = t.count("."), t.count(",")
    if cham and phay:                                   # 1,234.5 hoặc 1.234,5 — dấu xuất hiện CUỐI là thập phân
        tp = "." if t.rfind(".") > t.rfind(",") else ","
        t = t.replace("," if tp == "." else ".", "").replace(tp, ".")
    elif cham + phay == 1:
        d = "." if cham else ","
        sau = t.split(d)[1]
        t = t.replace(d, ".") if (co_he_so or len(sau) != 3) else t.replace(d, "")
    elif cham + phay > 1:
        d = "." if cham else ","
        nhom = t.split(d)
        if not all(len(g) == 3 for g in nhom[1:]):
            return None
        t = t.replace(d, "")
    try:
        return Decimal(t)
    except InvalidOperation:
        return None


def _don_vi_theo_chi_tieu(chi_tieu):
    c = kd(chi_tieu)
    if re.search(r"ty le|%|phan tram", c):
        return "pct", None
    if _TU_TIEN.search(c):
        return "dong", None
    if re.search(r"\bso luong\b|\bsl\b", c):
        return ("xe", None) if re.search(r"\bxe\b", c) else ("khac", "số lượng — log không ghi đơn vị")
    return "khac", "log không ghi đơn vị — bot tìm lỗi xác định theo mapping"


def doc_so(s, chi_tieu=""):
    """Ô số của log -> {gia_tri, don_vi[, ghi_chu]} theo $defs/so; ô trống hẳn -> None (không ghi trường)."""
    raw = str(s or "")
    t = raw.strip()
    if not t:
        return None
    dv_ct, ghi_ct = _don_vi_theo_chi_tieu(chi_tieu)
    if re.fullmatch(r"[-–—]+", t):
        return {"gia_tri": None, "don_vi": dv_ct, "ghi_chu": "ô '-' (dashboard/báo cáo để trống)"}
    am = t.startswith("(") and t.endswith(")")
    if am:
        t = t[1:-1].strip()
    if t.startswith("-") and re.match(r"-\s*\d", t):
        am, t = True, t.lstrip("-").strip()
    k = kd(t)
    don_vi = None
    if re.search(r"usd|\$", k):
        don_vi = "usd"
    elif "%" in k:
        don_vi = "pct"
    elif re.search(r"\bvnd\b|\bdong\b|\d\s*d\b|\bvnd", k):
        don_vi = "dong"
    elif re.search(r"\bxe\b", k):
        don_vi = "xe"
    he_so = Decimal(1)
    if re.search(r"\bty\b|\bti\b", k):
        he_so = Decimal(10) ** 9
    elif re.search(r"\btrieu\b|\btr\b", k):
        he_so = Decimal(10) ** 6
    elif re.search(r"\bnghin\b|\bngan\b", k):
        he_so = Decimal(1000)
    so = re.search(r"\d[\d.,\s]*\d|\d", k)
    v = _so_tu_chuoi(so.group(0).strip(), he_so != 1) if so else None
    if v is None:
        return {"gia_tri": None, "don_vi": "khac", "ghi_chu": f"không đọc được số: '{t}'"}
    v = v * he_so * (-1 if am else 1)
    v = int(v) if v == v.to_integral_value() else float(v)
    out = {"gia_tri": v, "don_vi": don_vi or dv_ct}
    if don_vi is None and ghi_ct:
        out["ghi_chu"] = f"{ghi_ct}; nguyên văn '{t}'"
    elif he_so != 1:
        out["ghi_chu"] = f"nguyên văn '{t}'"
    return out


def doc_dong_mapping(van_ban, sheet=None):
    """'dòng 10', 'dòng 22-29' trong văn bản -> ['<sheet>!10', '<sheet>!22', …] (giữ thứ tự, bỏ lặp)."""
    ra = []
    for a, b in re.findall(r"\bdong\s+(\d{1,4})(?:\s*(?:-|–|den|toi)\s*(\d{1,4}))?", kd(van_ban)):
        a, b = int(a), int(b or a)
        for n in range(a, min(b, a + 59) + 1) if b >= a else [a]:
            x = f"{sheet or '?'}!{n}"
            if x not in ra:
                ra.append(x)
    return ra


def _sheet_mapping(man):
    """Sheet mới nhất (đứng đầu danh sách) của tài liệu mapping có màn `man` trong mapping_index.json."""
    try:
        with open(MAPPING_INDEX, encoding="utf-8") as fh:
            idx = json.load(fh)
    except (OSError, ValueError):
        return None
    for tl in (idx.get("tai_lieu") or {}).values():
        if man in (tl.get("man") or []) and tl.get("sheet"):
            return tl["sheet"][0]
    return None


def _tim_man_mac_dinh(dashboard_me):
    from . import man_fe
    kq = man_fe.tim_man(dashboard_me)
    return kq[0]["man"] if kq else None


# ───────────────────────── cả khối ─────────────────────────
def _la_tieu_de(cells):
    k = [kd(c) for c in cells]
    return any(x.startswith("ten user") for x in k) or (any(x.startswith("chi tieu") for x in k)
                                                       and any(x.startswith("ky") for x in k))


def _cot_theo_tieu_de(cells):
    ra = {}
    for i, c in enumerate(cells):
        k = kd(c)
        for tien_to, khoa in _TEN_COT:
            if k.startswith(tien_to) and khoa not in ra.values():
                ra[i] = khoa
                break
    return ra


def tach_log(van_ban, tim_man=None, sheet_cua_man=None):
    """Khối TSV -> danh sách dòng đã chuẩn hoá (chưa tạo phiếu).

    Mỗi phần tử: {thu_tu, nguon_log, khoa_trung, trung_voi_thu_tu, lien_quan_thu_tu, canh_bao}. `thu_tu` = số thứ tự
    dòng dữ liệu trong khối dán (1-based, không tính tiêu đề) — cũng là `nguon_log.dong_excel` để bảng 'IT fix lỗi'
    trả về đúng thứ tự dòng người dùng dán."""
    tim_man = tim_man or _tim_man_mac_dinh
    sheet_cua_man = sheet_cua_man or _sheet_mapping
    rows = [r for r in csv.reader(io.StringIO(str(van_ban or "").strip("\n")), dialect="excel-tab")]
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        raise ValueError("khối log rỗng")
    cot = {i: k for i, k in enumerate(COT)}
    if _la_tieu_de(rows[0]):
        cot = _cot_theo_tieu_de(rows[0]) or cot
        rows = rows[1:]
    if len(rows) > MAX_DONG:
        raise ValueError(f"khối log có {len(rows)} dòng, vượt trần {MAX_DONG} dòng/lần — chia nhỏ")
    man_cache, truoc, ra, gap = {}, {}, [], {}
    for so_tt, r in enumerate(rows, 1):
        o = {k: (r[i] if i < len(r) else "") for i, k in cot.items()}
        for k in _KE_THUA:
            if o.get(k, "").strip():
                truoc[k] = o[k]
            else:
                o[k] = truoc.get(k, "")
        canh_bao = []
        dm = re.sub(r"\s+", " ", o.get("dashboard_me", "")).strip()
        if dm not in man_cache:
            man_cache[dm] = tim_man(dm) if dm else None
        man = man_cache[dm]
        if not man:
            canh_bao.append(f"không suy được mã màn từ '{dm}'")
        chi_tieu = re.sub(r"\s+", " ", o.get("chi_tieu", "")).strip()
        ky_raw = re.sub(r"\s+", " ", o.get("ky", "")).strip()
        ky = doc_ky(ky_raw)
        if not ky:
            canh_bao.append(f"không chuẩn hoá được kỳ '{ky_raw}'")
        mo_ta, pt = o.get("mo_ta", "").strip(), o.get("phan_tich", "").strip()
        nl = {"nguoi_bao": o.get("nguoi_bao", "").strip() or "?", "man": man or "?", "chi_tieu": chi_tieu or "?",
              "ky": ky or ky_raw or "?", "dong_excel": so_tt}
        if ky and ky != ky_raw:
            nl["ky_nguyen_van"] = ky_raw
        if dm:
            nl["dashboard_me"] = dm
        if o.get("chi_tieu_con", "").strip():
            nl["chi_tieu_con"] = re.sub(r"\s+", " ", o["chi_tieu_con"]).strip()
        if o.get("cost_center", "").strip():
            nl["khoi"] = o["cost_center"].strip()
        if o.get("stt", "").strip():
            nl["stt"] = o["stt"].strip()
        for k in ("so_bao_cao", "so_dashboard", "chenh_lech"):
            v = doc_so(o.get(k, ""), chi_tieu)
            if v is not None:
                nl[k] = v
        dmap = doc_dong_mapping(mo_ta + "\n" + pt, sheet_cua_man(man) if man else None)
        if dmap:
            nl["dong_mapping"] = dmap
        if mo_ta:
            nl["mo_ta"] = mo_ta
        if pt:
            nl["phan_tich_nghiep_vu"] = pt
        for k in ("it_fix", "trang_thai_log", "ghi_chu"):
            if o.get(k, "").strip():
                nl[k] = o[k].strip()
        khoa = (man or dm, kd(chi_tieu), ky or kd(ky_raw))
        trung = gap.get(khoa)
        if trung is None:
            gap[khoa] = so_tt
        lien = []
        m = re.search(r"giong (?:loi|dong)\s*(?:so\s*)?(\d+)", kd(mo_ta + " " + pt))
        if m:
            n = m[1]
            nl["lien_quan_ghi_chu"] = m.group(0)
            theo_stt = [x["thu_tu"] for x in ra if x["nguon_log"].get("stt") == n]
            lien = theo_stt or [x["thu_tu"] for x in ra if x["nguon_log"]["man"] == nl["man"]
                                and kd(x["nguon_log"]["chi_tieu"]) == kd(chi_tieu) and x["thu_tu"] != trung]
        ra.append({"thu_tu": so_tt, "nguon_log": nl, "khoa_trung": list(khoa), "trung_voi_thu_tu": trung,
                   "lien_quan_thu_tu": lien, "canh_bao": canh_bao})
    return ra

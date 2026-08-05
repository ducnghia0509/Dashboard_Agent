# -*- coding: utf-8 -*-
"""Deriver: BÁO CÁO NGÀY (HQKD theo từng ngày) — nguồn `received_reports/<FOLDER>/baocaohqkdngay/
B.<n>.<MÃ>.D.<YYYYMM>.<Tên>.xlsx` (ký tự **D** trong tên = Day; bản **M** là báo cáo tháng đã có
pipeline riêng, KHÔNG đụng tới).

PHẠM VI (chốt theo Mapping_Dashboard_QTTC.xlsx — cột "Đường link lấy dữ liệu ngày tạm thời trên
EXCEL", ngay bên phải cột "Map màn hình"): CHỈ 4 đơn vị có ô này khác "ko có" mới lên được báo cáo
ngày, và CHỈ cụm chỉ tiêu P&L (doanh thu / giá vốn / chi phí / lợi nhuận). Mọi chỉ tiêu khác của
màn Công nợ · Tồn kho · Tài sản · Thuế · Dòng tiền ghi "ko có" -> KHÔNG dựng số ngày cho chúng.

report_type RIÊNG có hậu tố `_D` (HQKD_D / PNLT_D / CHIPHI_D) — KHÔNG ghi đè HQKD/PNLT/CHIPHI của
báo cáo tháng. LÝ DO BẮT BUỘC: dòng ngày nằm CÙNG dataset tháng (để bộ lọc "Khoảng ngày" của FE
vẫn resolve ra đúng dataset kỳ). Nếu dùng chung report_type thì chế độ Tháng — vốn KHÔNG gửi
from/to nên cộng hết mọi dòng trong dataset — sẽ cộng cả 31 dòng ngày lẫn dòng tháng = số gấp đôi.
Backend chọn `_D` hay bản tháng theo `grain` của request (xem app/metrics/repository.py::_rt).

Đơn vị & layout (verify trên file thật kỳ 2026-08):
  · SRVF  (`B.1.TC.TCKT.D.202608.BaocaoHQKD.xlsx`) — layout "srvf": mỗi ngày 1 sheet tên "{d}.{m}"
    (vd "1.8" = 01/08). Header dòng 1: Mục | Khối | Mã số | Chỉ tiêu | <Showroom> | % DT | …
    Mã A-series y hệt sheet tháng T{mm}BC (A100 doanh thu, A300 tổng CP, A310 giá vốn, A600 lợi
    nhuận, U302 LNST). KHÔNG có cột tổng (cột "CHI NHÁNH VINFAST HÀ NỘI" luôn = 0) -> chỉ ghi dòng
    theo cost center, tổng khối = Σ showroom (đúng nhánh cc_v của repository._per_file_resolved).
    ⚠ VỊ TRÍ CỘT LỆCH GIỮA CÁC SHEET (sheet "1.8" showroom đầu ở cột 5, "2.8"/"3.8" ở cột 6 vì
    thêm 1 cột "% DT") -> BẮT BUỘC dò cột theo TÊN header từng sheet, không hardcode index.
  · XANHVINHPHUC (`B.6.XVP.D.<YYYYMM>.Baocaohqkdngay.xlsx`) — layout "kqkd": sheet "01".."31" theo
    ngày. Header dòng 4: CHỈ TIÊU | MÃ SỐ | Tổng cộng | HO | Depot Phú Thọ | Depot Vĩnh Phúc |
    Depot Tuyên Quang. Có cột tổng -> ghi thêm dòng "trực tiếp" (không cost center) như bản tháng.
  · HTXXANHTUYENQUANG / HTXXANHVINHPHUC (`B.6.HTX_*.D.<YYYYMM>.Baocaotaichinhrieng.xlsx`) — cũng
    layout "kqkd" nhưng header ở dòng 6 và CHỈ 1 cột giá trị (không có cost center).
  · ANTAXI (`B.7.AAG.TCKT.**M**.<YYYYMM>.Baocaotaichinhrieng.xlsx`) — layout "antaxi": sheet
    "1".."31" theo ngày (+ sheet "TH" tổng hợp, bỏ qua). Header dòng 6: TT | CHỈ TIÊU | Tỉ lệ |
    Sơn Tây | Tỉ lệ | Thái Nguyên | Tỉ lệ | Tổng cộng | Xe thương quyền | … | Lũy kế tháng.
    ⚠ TÊN FILE ghi `.M.` (trùng hệt báo cáo THÁNG ở `baocaotaichinhrieng/`) — chỉ THƯ MỤC
    `baocaohqkdngay/` phân biệt được, xem `_period_of`/`_in_day_dir`.
    ⚠ Số mục La Mã ở cột TT riêng + lệch so với XVP -> bảng anchor riêng, xem `_ANTAXI_ANCHOR`.
  · ANKHACHSAN (`B.10.AAG.TCKT.D.<YYYYMM>.Baocaotaichinhrieng.xlsx`) — layout "anks", KHÁC HẲN
    4 layout trên: CHỈ 1 SHEET, mỗi NGÀY là 1 CỘT (không phải mỗi ngày 1 sheet). Header dòng 3:
    TT | Nội dung chi phí | ĐVT | Tổng cộng | 1 | 2 | … | 31 — số ngày nằm sẵn ở header, không
    suy từ tên sheet. Chỉ 1 cơ sở (Garden Sơn Tây), không cost center. Xem `_anks_all_days`.

Neo dòng chỉ tiêu của layout "kqkd" theo TIỀN TỐ SỐ LA MÃ / số mục đã chuẩn hoá bỏ dấu, KHÔNG theo
"Mã số" và KHÔNG theo địa chỉ ô cứng (C17/C28/C100… như ghi chú trong file mapping): mã số bị TRÙNG
(HTX có 2 dòng mã "10": "1. Doanh thu bán hàng" và "3. Doanh thu thuần") và số thứ tự dòng LỆCH
giữa XVP (dòng 20) với HTX (dòng 22).

Tổng chi phí = Σ 6 cấu phần (IV giá vốn + VI biến đổi + VIII cố định + IX.2 tài chính + X.2 khác +
XII phân bổ chung) — verify khớp ĐÚNG dòng "18. Tổng chi phí" có sẵn trong file XVP ngày 01/08
(1.026.059.132 = 850.225.145 + 40.124.348 + 0 + 122.161.253 + 0 + 13.548.387). Lấy Σ cấu phần thay
vì dòng 18.1 để Σ lát CHIPHI_D luôn == tổng chi phí HQKD_D (file HTX không có dòng 18.1).

Chạy (dry-run, in tổng theo ngày, KHÔNG ghi):
  .venv/bin/python scripts/derive_hqkd_ngay.py <file.xlsx>
Ghi thật:
  .venv/bin/python scripts/derive_hqkd_ngay.py <file.xlsx> --write
"""
import argparse
import json
import os
import re
import sys
import unicodedata

import openpyxl
import psycopg

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [p for p in (os.path.dirname(_HERE), _HERE) if p not in sys.path]
from servers.common import source_catalog as _SC  # noqa: E402

DB_URL = (os.environ.get("DATABASE_URL") or os.environ.get("TC_DATABASE_URL")
          or "postgresql://tc:tc@localhost:5433/tc_dashboard")

RT_HQKD, RT_PNLT, RT_CHIPHI, RT_DTHU = "HQKD_D", "PNLT_D", "CHIPHI_D", "DTHU_D"
REPORT_TYPES = (RT_HQKD, RT_PNLT, RT_CHIPHI, RT_DTHU)

# Mã chỉ tiêu 01_HQKD (khớp app/metrics/repository.py: HQKD_REVENUE/COST/PROFIT_AT).
MA_DT, MA_CP, MA_LNTT = "1000", "1047", "1112"

_UNITS = {
    "SRVF": {"layout": "srvf", "cong_ty": "TC", "khoi": "Khối KD Vinfast - Showroom"},
    "XANHVINHPHUC": {"layout": "kqkd", "cong_ty": "XVP", "khoi": "Khối KD Vận tải Taxi Xanh"},
    "HTXXANHTUYENQUANG": {"layout": "kqkd", "cong_ty": "HTX_XTQ", "khoi": "Khối KD Vận tải Taxi Xanh"},
    "HTXXANHVINHPHUC": {"layout": "kqkd", "cong_ty": "HTX_XVP", "khoi": "Khối KD Vận tải Taxi Xanh"},
    "ANTAXI": {"layout": "antaxi", "cong_ty": "AAG", "khoi": "Khối KD Dịch vụ An Taxi"},
    "ANKHACHSAN": {"layout": "anks", "cong_ty": "AAG", "khoi": "Khối KD Dịch vụ An KS"},
}

# Cost center theo TỪ KHOÁ trong tên cột (dò theo tên, không theo vị trí — xem docstring).
# Mã CC lấy y hệt bản tháng (agent_cli._SR_SHOWROOM_CC / raw_rows thật của XVP).
_CC_SRVF = [("uong bi", "UB_SR"), ("b2b", "B2B_SR"), ("oceanpark", "OCP_SR"), ("long bien", "LB_SR"),
            ("smart city", "SMC_SR"), ("ha long", "HL_SR"), ("cam pha", "CP_SR"),
            ("vinh phuc", "VP_SR"), ("son tay", "ST_SR"), ("xuan mai", "XM_SR")]
_CC_XVP = [("depot phu tho", "PT_DP"), ("depot vinh phuc", "VP_DP"), ("depot tuyen quang", "TQ_DP"),
           ("ho", "HO_XVP")]
# An Taxi: header ghi tên tỉnh trần ("Sơn Tây" / "Thái Nguyên"), mã CC lấy y hệt master_data
# (app/data/master_data.json: ST_AT 'Depot Sơn Tây', TN_AT 'Depot Thái Nguyên').
_CC_ANTAXI = [("son tay", "ST_AT"), ("thai nguyen", "TN_AT")]

# Showroom Uông Bí thuộc pháp nhân VFQN (SRVF là thư mục ĐA-pháp-nhân) — giống bản tháng.
_CC_CONGTY = {"UB_SR": "VFQN"}


def _nd(s):
    """Chuẩn hoá: bỏ dấu, thường hoá, gộp khoảng trắng (nguồn hay gõ sai dấu: 'Gíá vốn', 'LỢI NHUÂN')."""
    s = str(s or "").strip().lower().replace("đ", "d")
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s)


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _source_id(path):
    folder = (_SC.raw_company_from_path(path) or "").upper()
    return f"{folder}::{os.path.basename(path)}"


_DAY_DIR = "baocaohqkdngay"


def _in_day_dir(path):
    """File có nằm trong thư mục `baocaohqkdngay/` không."""
    return _nd(os.path.basename(os.path.dirname(os.path.abspath(path)))).replace(" ", "") == _DAY_DIR


def _period_of(file_name, in_day_dir=False):
    """'B.1.TC.TCKT.D.202608.BaocaoHQKD.xlsx' -> '2026-08'. None nếu không phải file kỳ .D.<YYYYMM>.

    An Taxi gửi báo cáo NGÀY nhưng đặt tên '.M.<YYYYMM>' Y HỆT báo cáo THÁNG
    ('B.7.AAG.TCKT.M.202608.Baocaotaichinhrieng.xlsx' — cùng tên với file trong
    `baocaotaichinhrieng/`) -> chỉ THƯ MỤC phân biệt được. Vì vậy '.M.' CHỈ được coi là kỳ báo
    cáo ngày khi `in_day_dir`; nếu nới lỏng theo tên file thì báo cáo THÁNG của An Taxi sẽ bị
    gate ngày bắt và pipeline tháng không bao giờ chạy."""
    m = re.search(r"\.D\.(\d{4})(\d{2})", file_name)
    if m is None and in_day_dir:
        m = re.search(r"\.M\.(\d{4})(\d{2})", file_name)
    return f"{m.group(1)}-{m.group(2)}" if m else None


def is_daily_report(path):
    """File này có phải BÁO CÁO NGÀY của đơn vị đã cấu hình không (dùng cho gate ở agent_cli)."""
    folder = _source_id(path).split("::", 1)[0]
    return bool(_UNITS.get(folder)) and bool(_period_of(os.path.basename(path), _in_day_dir(path)))


# ---------------------------------------------------------------------------------------------
# Layout "srvf" — mã A-series, mỗi ngày 1 sheet "{d}.{m}"
# ---------------------------------------------------------------------------------------------
# Cấu phần TRỰC TIẾP của A300 (bản tháng parse từ công thức A300; workbook data_only KHÔNG còn
# công thức nên dùng thẳng danh sách fallback y hệt agent_cli._chiphi_recs_srvf).
_SRVF_CP_CODES = ["A310", "A320", "A325", "A330", "A340", "A350", "A360", "A500"]


def _srvf_day_sheets(wb, period):
    """[(sheet_name, 'YYYY-MM-DD')] — chỉ sheet tên '{ngày}.{tháng}' khớp THÁNG của kỳ."""
    y, mm = int(period[:4]), int(period[5:7])
    out = []
    for s in wb.sheetnames:
        m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})", s.strip())
        if m and int(m.group(2)) == mm and 1 <= int(m.group(1)) <= 31:
            out.append((s, f"{y:04d}-{mm:02d}-{int(m.group(1)):02d}"))
    return sorted(out, key=lambda x: x[1])


def _srvf_facts(rows):
    """rows -> [(cost_center, report_type, dim1, dim3, value_VND)] cho 1 ngày. [] nếu sai layout."""
    hdr = next((r for r in rows[:8] if any(_nd(c) == "ma so" for c in r if c is not None)), None)
    if hdr is None:
        return []
    ma_j = next(j for j, c in enumerate(hdr) if _nd(c) == "ma so")
    ten_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("chi tieu")), ma_j + 1)
    cols = [(cc, j) for kw, cc in _CC_SRVF
            for j in [next((j for j, c in enumerate(hdr) if isinstance(c, str) and kw in _nd(c)), None)]
            if j is not None]
    if len(cols) < 8:      # thiếu showroom = layout lạ -> không đoán
        return []
    by_code = {}
    for r in rows:
        c = str(r[ma_j]).strip().upper() if ma_j < len(r) and r[ma_j] not in (None, "") else ""
        if c and c not in by_code:
            by_code[c] = r

    def val(code, j):
        r = by_code.get(code)
        return _num(r[j]) if r is not None and j < len(r) else None

    from extract_chiphi import _nhom_cp
    facts = []
    for cc, j in cols:
        for code, rt, dim1 in (("A100", RT_HQKD, MA_DT), ("A300", RT_HQKD, MA_CP),
                               ("A600", RT_HQKD, MA_LNTT), ("A310", RT_PNLT, "Giá vốn hàng bán"),
                               ("A100", RT_PNLT, "Doanh thu HH, DV"),
                               ("A600", RT_PNLT, "Lợi nhuận trước thuế"),
                               ("A100", RT_DTHU, "Doanh thu thuần")):
            v = val(code, j)
            if v:
                facts.append((cc, rt, dim1, dim1, v))
        lnst = val("U302", j)
        lnst = lnst if lnst else val("A600", j)
        if lnst:
            facts.append((cc, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lnst))
        for code in _SRVF_CP_CODES:
            v = val(code, j)
            if not v:
                continue
            r = by_code[code]
            ten = str(r[ten_j]).strip() if ten_j < len(r) and r[ten_j] not in (None, "") else code
            facts.append((cc, RT_CHIPHI, _nhom_cp(ten, code), ten, v))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "kqkd" — XVP / HTX, mỗi ngày 1 sheet "01".."31"
# ---------------------------------------------------------------------------------------------
# (khoá, tiền tố nhãn đã chuẩn hoá) — neo theo NHÃN, xem docstring đầu file.
_KQKD_ANCHOR = {
    "dt_thuan": "3. doanh thu thuan",
    "gia_von": "iv. gia von hang ban",
    "ln_gop": "v. loi nhuan gop",
    "cp_bien_doi": "vi. chi phi bien doi",
    "cp_co_dinh": "viii. chi phi co dinh",
    "cp_tai_chinh": "2. chi phi tai chinh",
    "cp_khac": "2. chi phi khac",
    "pb_chung": "xii. phan bo chi phi chung",
    "lntt": "xiii. loi nhuan truoc thue",
    "lnst": "xv. loi nhuan sau thue",
}
# Cấu phần tổng chi phí -> (nhóm CP chuẩn mực, nhãn hiển thị). Nhóm khớp bản THÁNG của cùng đơn vị
# (raw_rows CHIPHI của XVP chỉ có 4 nhóm: Giá vốn / Chi phí tài chính / Chi phí bán hàng / Chi phí
# QLDN) để biểu đồ cơ cấu chi phí đọc được như nhau ở cả 2 chế độ; nhãn gốc giữ ở dim3.
_KQKD_CP = [
    ("gia_von", "Giá vốn hàng bán", "Giá vốn hàng bán"),
    ("cp_bien_doi", "Chi phí bán hàng", "Chi phí biến đổi"),
    ("cp_co_dinh", "Chi phí QLDN", "Chi phí cố định"),
    ("cp_tai_chinh", "Chi phí tài chính", "Chi phí tài chính"),
    ("cp_khac", "Chi phí khác", "Chi phí khác"),
    ("pb_chung", "Chi phí QLDN", "Phân bổ chi phí chung"),
]


def _kqkd_day_sheets(wb, period):
    y, mm = int(period[:4]), int(period[5:7])
    out = []
    for s in wb.sheetnames:
        t = s.strip()
        if t.isdigit() and 1 <= int(t) <= 31:
            out.append((s, f"{y:04d}-{mm:02d}-{int(t):02d}"))
    return sorted(out, key=lambda x: x[1])


def _kqkd_scan(rows, ccs, anchors):
    """Dò header + cột giá trị + neo dòng chỉ tiêu cho layout dạng KQKD (dùng cho cả 'kqkd' và
    'antaxi'). -> (cols, anchored, val) ; cols=[] nếu sai layout."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None:
        return [], {}, None
    hdr = rows[hdr_i]
    ten_j = next(j for j, c in enumerate(hdr) if _nd(c) == "chi tieu")
    # Cột giá trị: ƯU TIÊN các cột cost center (Depot/HO/tỉnh). CỐ Ý BỎ cột "Tổng cộng" khi đã có
    # cột cost center — Σ cost center khớp ĐÚNG cột tổng (verify XVP 01/08: 283.756.022 +
    # 480.309.773 + 140.030.764 = 904.096.560; An Taxi 01/08 DT thuần: 76.225.625 + 34.226.640 =
    # 110.452.265), nên ghi thêm dòng tổng chỉ tạo nguy cơ đếm đôi ở các truy vấn KHÔNG đi qua
    # repository._per_file_resolved (daily_series/sum_by_dim1 — biểu đồ xu hướng ngày).
    # Đơn vị 1 cột (HTX, không có cost center) -> dùng cột tổng / cột ngay sau "MÃ SỐ".
    cols = [(cc, j) for j, c in enumerate(hdr) if j != ten_j
            for cc in [next((cc for kw, cc in ccs if _nd(c) == kw), None)] if cc]
    if not cols:
        tong_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("tong cong")), None)
        if tong_j is None:
            ma_j = next((j for j, c in enumerate(hdr) if _nd(c) == "ma so"), ten_j)
            tong_j = ma_j + 1
        cols = [(None, tong_j)]

    anchored = {}
    for r in rows[hdr_i + 1:]:
        if not r or ten_j >= len(r):
            continue
        n = _nd(r[ten_j])
        for key, pref in anchors.items():
            if key not in anchored and n.startswith(pref):
                anchored[key] = r

    def val(key, j):
        r = anchored.get(key)
        return _num(r[j]) if r is not None and j < len(r) else None

    return cols, anchored, val


def _kqkd_facts(rows):
    """rows -> [(cost_center|None, report_type, dim1, dim3, value_VND)]. [] nếu sai layout."""
    cols, anchored, val = _kqkd_scan(rows, _CC_XVP, _KQKD_ANCHOR)
    if not cols or "dt_thuan" not in anchored:
        return []

    facts = []
    for cc, j in cols:
        dt = val("dt_thuan", j)
        if dt:
            facts.append((cc, RT_HQKD, MA_DT, MA_DT, dt))
            facts.append((cc, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
            facts.append((cc, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
        tong_cp = 0.0
        for key, nhom, ten in _KQKD_CP:
            v = val(key, j)
            if not v:
                continue
            tong_cp += v
            facts.append((cc, RT_CHIPHI, nhom, ten, v))
        if tong_cp:
            facts.append((cc, RT_HQKD, MA_CP, MA_CP, tong_cp))
        for key, rt, dim1 in (("lntt", RT_HQKD, MA_LNTT), ("gia_von", RT_PNLT, "Giá vốn hàng bán"),
                              ("ln_gop", RT_PNLT, "Lợi nhuận gộp"),
                              ("lntt", RT_PNLT, "Lợi nhuận trước thuế"),
                              ("lnst", RT_PNLT, "Lợi nhuận sau thuế")):
            v = val(key, j)
            if v:
                facts.append((cc, rt, dim1, dim1, v))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "antaxi" — An Taxi (ANTAXI/baocaohqkdngay), mỗi ngày 1 sheet "1".."31" + sheet "TH"
# ---------------------------------------------------------------------------------------------
# KHÁC layout "kqkd" ở 2 điểm (verify file thật B.7.AAG.TCKT.M.202608, sheet 1-31, header dòng 6):
#  1. Số mục La Mã nằm ở cột "TT" RIÊNG, KHÔNG dính vào nhãn: dòng 33 = TT "III" + CHỈ TIÊU
#     "DOANH THU THUẦN" (XVP ghi "3. Doanh thu thuần" trong CÙNG 1 ô) -> neo theo NHÃN TRẦN.
#  2. Số La Mã của An Taxi lệch hẳn XVP (LNTT = XI ở An Taxi vs XIII ở XVP; LNST = XIII vs XV;
#     An Taxi KHÔNG có mục "phân bổ chi phí chung") -> phải có bảng anchor riêng, không dùng chung.
# Đã verify nhãn nào cũng khớp DUY NHẤT 1 dòng trên cả 31 sheet (các dòng con "Giá vốn bán xe…",
# "Chi phí hoạt động khác", "LỢI NHUẬN TRƯỚC KHẤU HAO…(EBITDA)" đều KHÔNG startswith nhãn neo).
_ANTAXI_ANCHOR = {
    "dt_gross": "doanh thu ban hang",              # I  (mã 100) — DT GỘP
    "giam_tru": "cac khoan giam tru doanh thu",    # II (mã 110)
    "dt_thuan": "doanh thu thuan",                 # III (mã 120)
    "gia_von": "gia von hang ban",                 # IV (mã 130)
    "ln_gop": "lai gop",                           # V  (mã 140) — An Taxi ghi "LÃI GỘP"
    "cp_bien_doi": "chi phi bien doi",             # VI (mã 150)
    "cp_co_dinh": "chi phi co dinh",               # VIII (mã 170)
    "dt_tai_chinh": "doanh thu tai chinh",         # IX.1
    "cp_tai_chinh": "chi phi tai chinh",           # IX.2 (mã 182)
    "tn_khac": "thu nhap khac",                    # X.1
    "cp_khac": "chi phi khac",                     # X.2 (mã 192)
    "lntt": "loi nhuan truoc thue",                # XI (mã 200)
    "lnst": "loi nhuan sau thue",                  # XIII (mã 220)
}
# Tổng chi phí = IV + VI + VIII + IX.2 + X.2 (đúng công thức cột "Công thức" của mapping An Taxi;
# file KHÔNG in dòng tổng chi phí nào để đối chiếu). Verify vòng kín theo LNTT có sẵn trong file —
# 01/08 Sơn Tây: 76.225.625 (DTT) − 93.342.155 (Σ 5 cấu phần) + 2.965.749 (thu nhập khác) + 0
# (DT tài chính) = −14.150.781 = ĐÚNG dòng "XI. LỢI NHUÂN TRƯỚC THUẾ TNDN (EBT)".
# dim1 (nhóm CP) lấy Y HỆT bản THÁNG của CHÍNH An Taxi — raw_rows CHIPHI T6 có đúng 5 nhóm:
# Giá vốn hàng bán / Chi phí biến đổi / Chi phí cố định / Chi phí tài chính / Chi phí khác
# (KHÁC XVP dùng 'Chi phí bán hàng'/'Chi phí QLDN') để biểu đồ cơ cấu chi phí đọc được như nhau
# ở cả 2 chế độ Ngày/Tháng.
_ANTAXI_CP = [
    ("gia_von", "Giá vốn hàng bán", "Giá vốn hàng bán"),
    ("cp_bien_doi", "Chi phí biến đổi", "Chi phí biến đổi"),
    ("cp_co_dinh", "Chi phí cố định", "Chi phí cố định"),
    ("cp_tai_chinh", "Chi phí tài chính", "Chi phí tài chính"),
    ("cp_khac", "Chi phí khác", "Chi phí khác"),
]


def _antaxi_facts(rows):
    """rows -> [(cost_center, report_type, dim1, dim3, value_VND)]. [] nếu sai layout."""
    cols, anchored, val = _kqkd_scan(rows, _CC_ANTAXI, _ANTAXI_ANCHOR)
    if not cols or "dt_thuan" not in anchored:
        return []

    facts = []
    for cc, j in cols:
        dt = val("dt_thuan", j)
        if dt:
            facts.append((cc, RT_HQKD, MA_DT, MA_DT, dt))
            facts.append((cc, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
        tong_cp = 0.0
        for key, nhom, ten in _ANTAXI_CP:
            v = val(key, j)
            if not v:
                continue
            tong_cp += v
            facts.append((cc, RT_CHIPHI, nhom, ten, v))
        if tong_cp:
            facts.append((cc, RT_HQKD, MA_CP, MA_CP, tong_cp))
        # PNLT: giữ ĐÚNG bộ dim1 bản THÁNG của An Taxi. Lưu ý An Taxi là đơn vị DUY NHẤT có dòng
        # giảm trừ tách riêng, và bản tháng ghi CẢ HAI tên cho DT GỘP ("Doanh thu bán hàng và cung
        # cấp dịch vụ" = tên đích danh cho bảng "Cấu trúc Doanh thu" ở revenue.py::_gross_of, và
        # "Doanh thu HH, DV" = tên chung cho KPI #1 ở overview.py) — cùng 1 giá trị, KHÔNG cộng đôi
        # vì 2 truy vấn lọc dim1_ilike khác nhau. Bỏ 1 trong 2 là mất số ở đúng 1 trong 2 chỗ đó.
        # ⚠ "Doanh thu HH, DV" của An Taxi là DT GỘP (mã 100), KHÁC XVP/HTX (= DT thuần).
        for key, rt, dim1 in (("lntt", RT_HQKD, MA_LNTT),
                              ("dt_gross", RT_PNLT, "Doanh thu bán hàng và cung cấp dịch vụ"),
                              ("dt_gross", RT_PNLT, "Doanh thu HH, DV"),
                              ("giam_tru", RT_PNLT, "Các khoản giảm trừ doanh thu"),
                              ("gia_von", RT_PNLT, "Giá vốn hàng bán"),
                              ("ln_gop", RT_PNLT, "Lợi nhuận gộp"),
                              ("dt_tai_chinh", RT_PNLT, "Doanh thu tài chính"),
                              ("tn_khac", RT_PNLT, "Thu nhập khác"),
                              ("lntt", RT_PNLT, "Lợi nhuận trước thuế"),
                              ("lnst", RT_PNLT, "Lợi nhuận sau thuế")):
            v = val(key, j)
            if v:
                facts.append((cc, rt, dim1, dim1, v))
    return facts


_FACTS_FN = {"srvf": _srvf_facts, "kqkd": _kqkd_facts, "antaxi": _antaxi_facts}


# ---------------------------------------------------------------------------------------------
# Layout "anks" — An KS (ANKHACHSAN), KHÁC HẲN 3 layout trên: chỉ 1 SHEET DUY NHẤT, mỗi NGÀY là
# 1 CỘT (không phải mỗi ngày 1 sheet). Header dòng 3: TT | Nội dung chi phí | ĐVT | Tổng cộng |
# 1 | 2 | … | 31 (số ngày nằm NGAY Ở HEADER, không cần suy). Nhãn chỉ tiêu ở cột "Nội dung chi
# phí", neo theo TIỀN TỐ đã chuẩn hoá (mã La Mã I/II/III ở cột riêng bên trái, không dính nhãn).
# CỘT "Tổng cộng" = Σ các cột ngày trong tháng (verify ngày 1-4/08: 2.975.000+1.600.000+1.615.000
# +1.400.000 = 7.590.000 = đúng ô Tổng cộng dòng I) -> BỎ cột này, chỉ lấy cột NGÀY để không đếm
# đôi khi gộp nhiều ngày (giống lý do bỏ cột "Tổng cộng" ở layout kqkd).
# KHÔNG cost center: An KS chỉ 1 cơ sở (Garden Sơn Tây) — verify raw_rows bản THÁNG: cột
# cost_center của khối An KS luôn NULL ở mọi report_type/mọi kỳ.
# dim1 PNLT/CHIPHI lấy Y HỆT bản THÁNG của An KS (rà DB 2026-06):
#   CHIPHI: 'Giá vốn hàng bán' / 'Chi phí chung' / 'Chi phí lương + CP khác cho CNV' (3 nhóm,
#     KHÁC An Taxi/XVP) — đúng 3 dòng con II.1/II.2/II.3 của file.
#   PNLT: ghi CẢ HAI tên gross ('Doanh thu HH, DV' VÀ 'Doanh thu bán hàng và cung cấp dịch vụ',
#     cùng giá trị — như An Taxi, xem doanhthu-cautructructure-gross-giamtru) + 'Giá vốn hàng
#     bán' + CHỈ 'Lợi nhuận sau thuế' (An KS KHÔNG có dòng thuế TNDN -> LNTT=LNST, bản THÁNG
#     không ghi dim1 'Lợi nhuận trước thuế' riêng, xem An.xlsx spec — daily giữ đúng convention,
#     KHÔNG tự thêm dòng monthly không có). 'Lợi nhuận gộp' KHÔNG ghi tường minh — bản THÁNG lưu
#     nó nhưng do BACKEND tự tính (revenue.py::pnl.py fallback DT-giá vốn theo khối khi thiếu
#     dòng riêng), daily để trống cho backend tự suy, tránh trùng 2 nguồn.
_ANKS_ANCHOR = {
    "dt": "tong doanh thu",              # I
    "tong_cp": "tong chi phi",           # II (nhãn gốc có dấu ':' cuối -> startswith vẫn khớp)
    "gia_von": "chi phi gia von",        # II.1
    "chi_chung": "chi phi chung",        # II.2
    "luong": "chi phi luong",            # II.3 "CHI PHÍ LƯƠNG + CP KHÁC CHO CNV"
    "lntt": "loi nhuan (i-ii)",          # III — An KS: LNTT = LNST (không có dòng thuế TNDN)
}


def _anks_all_days(rows, period):
    """1 sheet, mỗi cột 1 ngày -> [(ngay, facts), ...] cho TRỌN THÁNG trong 1 lần gọi (khác hẳn
    3 layout trên vốn mỗi ngày gọi facts_fn riêng — xem docstring khối trên)."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "tong cong" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    tong_j = next(j for j, c in enumerate(hdr) if _nd(c) == "tong cong")
    y, mm = int(period[:4]), int(period[5:7])
    day_cols = [(j, int(c)) for j, c in enumerate(hdr) if j > tong_j and isinstance(c, int)
                and not isinstance(c, bool) and 1 <= c <= 31]

    label_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("noi dung")), 2)
    anchored = {}
    for r in rows[hdr_i + 1:]:
        if not r or label_j >= len(r):
            continue
        n = _nd(r[label_j])
        for key, pref in _ANKS_ANCHOR.items():
            if key not in anchored and n.startswith(pref):
                anchored[key] = r

    def val(key, j):
        r = anchored.get(key)
        return _num(r[j]) if r is not None and j < len(r) else None

    per_day = []
    for j, d in day_cols:
        facts = []
        dt = val("dt", j)
        if dt:
            facts.append((None, RT_HQKD, MA_DT, MA_DT, dt))
            facts.append((None, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
            facts.append((None, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
            facts.append((None, RT_PNLT, "Doanh thu bán hàng và cung cấp dịch vụ",
                          "Doanh thu bán hàng và cung cấp dịch vụ", dt))
        tong_cp = val("tong_cp", j)
        if tong_cp:
            facts.append((None, RT_HQKD, MA_CP, MA_CP, tong_cp))
        for key, dim1 in (("gia_von", "Giá vốn hàng bán"), ("chi_chung", "Chi phí chung"),
                          ("luong", "Chi phí lương + CP khác cho CNV")):
            v = val(key, j)
            if v:
                facts.append((None, RT_CHIPHI, dim1, dim1, v))
        gia_von = val("gia_von", j)
        if gia_von:
            facts.append((None, RT_PNLT, "Giá vốn hàng bán", "Giá vốn hàng bán", gia_von))
        lntt = val("lntt", j)
        if lntt:
            facts.append((None, RT_HQKD, MA_LNTT, MA_LNTT, lntt))
            facts.append((None, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lntt))
        if facts:
            per_day.append((f"{y:04d}-{mm:02d}-{d:02d}", facts))
    return per_day


# ---------------------------------------------------------------------------------------------
def derive(path, write=False):
    folder = _source_id(path).split("::", 1)[0]
    unit = _UNITS.get(folder)
    period = _period_of(os.path.basename(path), _in_day_dir(path))
    if not unit or not period:
        return {"ok": False, "skip": True}

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if unit["layout"] == "anks":
            rows = [list(r) for r in wb[wb.sheetnames[0]].iter_rows(values_only=True)]
            per_day = _anks_all_days(rows, period)
        else:
            srvf = unit["layout"] == "srvf"
            facts_fn = _FACTS_FN[unit["layout"]]
            sheets = _srvf_day_sheets(wb, period) if srvf else _kqkd_day_sheets(wb, period)
            per_day = []
            for sheet, ngay in sheets:
                rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
                facts = facts_fn(rows)
                if facts:
                    per_day.append((ngay, facts))
    finally:
        wb.close()

    if not per_day:
        return {"ok": False, "error": f"không đọc được sheet ngày nào (kỳ {period}, layout {unit['layout']})"}

    out = {"ok": True, "file": os.path.basename(path), "period": period, "cong_ty": unit["cong_ty"],
           "layout": unit["layout"], "days": len(per_day),
           "tong_theo_ngay": {
               ngay: {"doanh_thu_ty": round(sum(v for _, rt, d1, _, v in f
                                                if rt == RT_HQKD and d1 == MA_DT) * 1e-9, 9),
                      "chi_phi_ty": round(sum(v for _, rt, d1, _, v in f
                                              if rt == RT_HQKD and d1 == MA_CP) * 1e-9, 9),
                      "lntt_ty": round(sum(v for _, rt, d1, _, v in f
                                           if rt == RT_HQKD and d1 == MA_LNTT) * 1e-9, 9)}
               for ngay, f in per_day}}
    if not write:
        return out

    source_file = _source_id(path)
    conn = psycopg.connect(DB_URL)
    try:
        cur = conn.cursor()
        cur.execute("SELECT id FROM datasets WHERE kind='month' AND period=%s "
                    "ORDER BY created_at DESC LIMIT 1", (period,))
        row = cur.fetchone()
        if not row:
            out["ok"] = False
            out["error"] = f"chưa có dataset tháng {period} — nạp báo cáo THÁNG trước rồi chạy lại"
            return out
        dataset_id = row[0]
        # idempotent: xoá bản cũ CÙNG source_file (mỗi file phủ trọn 1 tháng ngày).
        cur.execute("DELETE FROM raw_rows WHERE report_type = ANY(%s) AND source_file=%s",
                    (list(REPORT_TYPES), source_file))
        payload = json.dumps({"unit": "ty", "grain": "day"}, ensure_ascii=False)
        recs, i = [], 0
        for ngay, facts in per_day:
            for cc, rt, dim1, dim3, v in facts:
                i += 1
                recs.append((dataset_id, rt, 6200000 + i, ngay,
                             _CC_CONGTY.get(cc) or unit["cong_ty"], unit["khoi"], cc, period,
                             round(v * 1e-9, 9), None, dim1, None, dim3, payload, source_file))
        cur.executemany(
            "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
            "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, source_file) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", recs)
        conn.commit()
        out["written"] = len(recs)
    finally:
        conn.close()
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--write", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    print(json.dumps(derive(a.file, a.write), ensure_ascii=False, indent=2))

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
}

# Cost center theo TỪ KHOÁ trong tên cột (dò theo tên, không theo vị trí — xem docstring).
# Mã CC lấy y hệt bản tháng (agent_cli._SR_SHOWROOM_CC / raw_rows thật của XVP).
_CC_SRVF = [("uong bi", "UB_SR"), ("b2b", "B2B_SR"), ("oceanpark", "OCP_SR"), ("long bien", "LB_SR"),
            ("smart city", "SMC_SR"), ("ha long", "HL_SR"), ("cam pha", "CP_SR"),
            ("vinh phuc", "VP_SR"), ("son tay", "ST_SR"), ("xuan mai", "XM_SR")]
_CC_XVP = [("depot phu tho", "PT_DP"), ("depot vinh phuc", "VP_DP"), ("depot tuyen quang", "TQ_DP"),
           ("ho", "HO_XVP")]

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


def _period_of(file_name):
    """'B.1.TC.TCKT.D.202608.BaocaoHQKD.xlsx' -> '2026-08'. None nếu không phải file kỳ .D.<YYYYMM>."""
    m = re.search(r"\.D\.(\d{4})(\d{2})", file_name)
    return f"{m.group(1)}-{m.group(2)}" if m else None


def is_daily_report(path):
    """File này có phải BÁO CÁO NGÀY của đơn vị đã cấu hình không (dùng cho gate ở agent_cli)."""
    folder = _source_id(path).split("::", 1)[0]
    return bool(_UNITS.get(folder)) and bool(_period_of(os.path.basename(path)))


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


def _kqkd_facts(rows):
    """rows -> [(cost_center|None, report_type, dim1, dim3, value_VND)]. [] nếu sai layout."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    ten_j = next(j for j, c in enumerate(hdr) if _nd(c) == "chi tieu")
    # Cột giá trị: ƯU TIÊN các cột cost center (Depot/HO). CỐ Ý BỎ cột "Tổng cộng" khi đã có cột
    # cost center — Σ depot khớp ĐÚNG cột tổng (verify XVP 01/08: 283.756.022 + 480.309.773 +
    # 140.030.764 = 904.096.560), nên ghi thêm dòng tổng chỉ tạo nguy cơ đếm đôi ở các truy vấn
    # KHÔNG đi qua repository._per_file_resolved (daily_series/sum_by_dim1 — biểu đồ xu hướng ngày).
    # Đơn vị 1 cột (HTX, không có cost center) -> dùng cột tổng / cột ngay sau "MÃ SỐ".
    cols = [(cc, j) for j, c in enumerate(hdr) if j != ten_j
            for cc in [next((cc for kw, cc in _CC_XVP if _nd(c) == kw), None)] if cc]
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
        for key, pref in _KQKD_ANCHOR.items():
            if key not in anchored and n.startswith(pref):
                anchored[key] = r
    if "dt_thuan" not in anchored:
        return []

    def val(key, j):
        r = anchored.get(key)
        return _num(r[j]) if r is not None and j < len(r) else None

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
def derive(path, write=False):
    folder = _source_id(path).split("::", 1)[0]
    unit = _UNITS.get(folder)
    period = _period_of(os.path.basename(path))
    if not unit or not period:
        return {"ok": False, "skip": True}

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        srvf = unit["layout"] == "srvf"
        sheets = _srvf_day_sheets(wb, period) if srvf else _kqkd_day_sheets(wb, period)
        per_day = []
        for sheet, ngay in sheets:
            rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
            facts = _srvf_facts(rows) if srvf else _kqkd_facts(rows)
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

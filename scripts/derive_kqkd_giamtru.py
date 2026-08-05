# -*- coding: utf-8 -*-
"""Deriver BỔ SUNG: breakdown "Giảm trừ doanh thu" (Chiết khấu thương mại/TK 5211, Hàng bán bị
trả lại/TK 5212, Giảm giá hàng bán/TK 5213, phần dư -> "Các khoản giảm trừ khác") cho bảng
"Cấu trúc Doanh thu" màn Doanh thu – Giá vốn (spec user 2026-07-31).

KHÔNG ghi lại "Doanh thu bán hàng & CCDV" (gross, mã 01): rà DB thấy field này ĐÃ CÓ SẴN rộng khắp
qua dim1 "Doanh thu HH, DV" (mọi deriver KQKD hiện có đều ghi — xem revenue.py::_gross_of), viết
lại sẽ tạo dữ liệu trùng/xung đột. Deriver này CHỈ bịt lỗ hổng "giảm trừ":

  · Xanh VP (+ 2 HTX cùng mẫu): sheet 'HQKD' có SẴN 3 dòng "2.1. Chiết khấu thương mại" /
    "2.2 Hàng bán bị trả lại" / "2.3 Giảm giá hàng bán" — đọc trực tiếp CÙNG cột kỳ này (method A).
  · An Taxi, GA: sheet CĐPS/TC_CDPS có mở TK 5211 (Chiết khấu thương mại) — dò MÃ TK ở BẤT KỲ
    sheet nào trong file (method B), KHÔNG cứng tên sheet (đơn vị chưa từng phát sinh 5212/5213
    thì bảng TK không có dòng đó -> coi = 0, không suy diễn).
  · Đơn vị không có sheet KQKD/HQKD nhận diện được (HO quản trị 511*, SRVF P&L mã A-series, Hưng
    Thịnh hợp nhất công thức hỏng #REF!) -> return None, giữ "—" (không suy diễn).

CỘT KỲ NÀY: dò theo NHÃN HEADER — 'Năm nay'/'Kỳ này' (mẫu TT200 B02-DN chuẩn: CHỈ TIÊU|Mã số|
Thuyết minh|Năm nay|Năm trước, ĐÃ XÁC MINH giống hệt ở An Taxi/An KS/GA/Dự án/Trạm sạc) hoặc
'T<tháng>.<năm>' vd 'T5.2026' (mẫu quản trị Xanh VP + 2 HTX — CHỈ TIÊU|Mã số|Thuyết minh|T5.2026|
%DT; đã sửa 2026-08-01, bản đầu chỉ nhận 'TỔNG CỘNG' nên luôn None ở 3 đơn vị này — xem
_PERIOD_MONTH_RE). KHÔNG neo qua DB (thử trước, phát hiện DTHU lưu theo TỪNG cost-center/company
khác cấu trúc 1-cột-1-file nên so lệch giả — bỏ, dùng thẳng nhãn cột đáng tin hơn).

GHI: report_type PNLT, pattern additive (như PTRA_ADV/derive_srvf_cdps) — KHÔNG qua import_filled.
  · 3 dòng chi tiết (Chiết khấu/Trả lại/Giảm giá) + "Các khoản giảm trừ khác": LUÔN ghi (kể cả 0)
    khi tìm được sheet + cột kỳ, để phân biệt "biết chắc = 0" với "không có dữ liệu" (None).
  · Dòng TỔNG "Giảm trừ doanh thu": CHỈ ghi khi (source_file, kỳ) CHƯA có sẵn dòng PNLT nào tên
    "Các khoản giảm trừ doanh thu"/"Giảm trừ doanh thu" (tránh đè/đếm đôi với An Taxi — nguồn khác,
    sheet BCQT PT, đã ghi từ trước qua deriver riêng, giá trị trùng khớp TK 5211 nên không cần ghi
    lại). Đơn vị đang "—" (An KS/GA/Dự án/Trạm sạc mã 02 = 0 mọi tháng, Xanh VP/2 HTX chưa từng có
    dòng nào) -> điền dòng TỔNG này, biến "—" cũ thành "0" tường minh — ĐÚNG hơn vì nguồn thực sự
    ghi 0, không phải thiếu dữ liệu.

Full precision (không round từng dòng — quy ước 2026-07-30).
Chạy: .venv/bin/python scripts/derive_kqkd_giamtru.py <file.xlsx> --period 2026-06
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers.common import be_bridge as bb  # noqa: E402
import agent_cli as A  # noqa: E402

norm = lambda v: bb.remove_diacritics("" if v is None else str(v)).strip().lower()  # noqa: E731
_EPS = 1000.0 * 1e-9   # 1.000đ theo đơn vị tỷ

DIM_CKTM = "Chiết khấu thương mại"
DIM_TRALAI = "Hàng bán bị trả lại"
DIM_GIAMGIA = "Giảm giá hàng bán"
DIM_KHAC = "Các khoản giảm trừ khác"
DIM_TONG = "Giảm trừ doanh thu"
_EXISTING_TONG_NAMES = ("Các khoản giảm trừ doanh thu", "Giảm trừ doanh thu")
_SUB_PHRASES = {DIM_CKTM: "chiet khau thuong mai", DIM_TRALAI: "hang ban bi tra lai",
                DIM_GIAMGIA: "giam gia hang ban"}
_TK_MAP = {"5211": DIM_CKTM, "5212": DIM_TRALAI, "5213": DIM_GIAMGIA}
_PERIOD_LABELS = ("nam nay", "ky nay", "tong cong")
_PERIOD_MONTH_RE = re.compile(r"^t\d{1,2}\.\d{4}$")  # mẫu Xanh VP/2 HTX: header cột = 'T5.2026'


def _num(v):
    return v if isinstance(v, (int, float)) else None


def _kqkd_sheet(wb):
    # Ưu tiên sheet TÊN CHÍNH XÁC 'HQKD' trước: Xanh VP (+2 HTX cùng mẫu) có CẢ 'KQKD' (tổng đơn
    # giản, KHÔNG breakdown) LẪN 'HQKD' (quản trị, CÓ dòng 2.1/2.2/2.3 chi tiết giảm trừ) trong
    # CÙNG file — vơ theo 'kqkd' chung sẽ vớ nhầm sheet thiếu breakdown dù số tổng vẫn tình cờ khớp.
    exact = next((s for s in wb.sheetnames if norm(s).strip() == "hqkd"), None)
    if exact:
        return exact
    for s in wb.sheetnames:
        n = norm(s).replace(" ", "")
        if "kqkd" in n or "ketqua" in norm(s) or "kqhdkd" in n:
            return s
    return None


def _period_col(rows, hdr_scan=12):
    for r in rows[:hdr_scan]:
        for j, c in enumerate(r):
            if not isinstance(c, str):
                continue
            n = norm(c).strip()
            if n in _PERIOD_LABELS or _PERIOD_MONTH_RE.match(n):
                return j
    return None


def _breakdown_same_sheet(rows, col_j):
    """Method A: dòng '2.1 Chiết khấu thương mại' / '2.2 Hàng bán bị trả lại' / '2.3 Giảm giá
    hàng bán' NẰM NGAY trong sheet KQKD (mẫu Xanh VP) — đọc value tại CÙNG cột đã xác định."""
    out = {}
    for r in rows:
        lbl = next((str(c) for c in r if isinstance(c, str) and len(c.strip()) > 2), None)
        if not lbl:
            continue
        n = norm(lbl)
        for dim, phrase in _SUB_PHRASES.items():
            if dim not in out and phrase in n:
                v = _num(r[col_j]) if col_j < len(r) else None
                out[dim] = (v or 0.0) * 1e-9
                break
    return out


def _breakdown_by_tk(wb):
    """Method B: dò MÃ TK 5211/5212/5213 ở BẤT KỲ sheet nào (CĐPS/TC_CDPS…) — giá trị = số lớn
    nhất (trị tuyệt đối) trong các ô số của dòng đó (TK trung gian tự cân Nợ=Có mỗi kỳ, Dư đầu/
    cuối thường = 0 -> max = đúng số phát sinh trong kỳ; hàng toàn 0 -> max = 0, vẫn đúng)."""
    out = {}
    for sn in wb.sheetnames:
        try:
            ws_rows = wb[sn].iter_rows(values_only=True)
        except Exception:  # noqa: BLE001
            continue
        for r in ws_rows:
            for j, c in enumerate(r):
                if not isinstance(c, str):
                    continue
                code = c.strip()
                if code.endswith(".0"):
                    code = code[:-2]
                dim = _TK_MAP.get(code)
                if dim is None or dim in out:
                    continue
                nums = [x for k, x in enumerate(r) if k != j and isinstance(x, (int, float))]
                v = max(nums, key=abs) if nums else 0.0
                out[dim] = v * 1e-9
        if len(out) == 3:
            break
    return out


def compute(path):
    """Trả dict {tong, ck_tm, tra_lai, giam_gia, khac} đơn vị TỶ, hoặc None nếu không nhận diện
    được sheet KQKD/HQKD hoặc không tìm được cột kỳ này."""
    wb = bb.fast_load_workbook(path, read_only=True, data_only=True)
    try:
        sn = _kqkd_sheet(wb)
        if not sn:
            return None
        rows = [list(r) for r in wb[sn].iter_rows(values_only=True)]
        r02 = next((r for r in rows
                    if "giam tru doanh thu" in norm(" ".join(str(c) for c in r if isinstance(c, str)))),
                   None)
        col_j = _period_col(rows)
        if r02 is None or col_j is None:
            return None
        tong = (_num(r02[col_j]) or 0.0) * 1e-9 if col_j < len(r02) else 0.0
        sub = _breakdown_same_sheet(rows, col_j)          # method A trước (cùng bảng, đáng tin hơn)
        for dim, v in _breakdown_by_tk(wb).items():        # method B bổ sung khoản A chưa thấy
            sub.setdefault(dim, v)
    finally:
        wb.close()
    ck_tm = sub.get(DIM_CKTM, 0.0)
    tra_lai = sub.get(DIM_TRALAI, 0.0)
    giam_gia = sub.get(DIM_GIAMGIA, 0.0)
    khac = tong - (ck_tm + tra_lai + giam_gia)
    if abs(khac) < _EPS:             # bụi float khi 3 khoản đã cộng đủ 100% mã 02
        khac = 0.0
    return {"tong": tong, "ck_tm": ck_tm, "tra_lai": tra_lai, "giam_gia": giam_gia, "khac": khac}


def extract(path, period, khoi=None, cong_ty=None):
    khoi = khoi or A._khoi_of(path)
    if not khoi:
        return {"ok": False, "skip": True, "reason": "không suy được khối"}
    r = compute(path)
    src = A._source_id(path)
    db = bb.db.get_db()
    db.execute("DELETE FROM raw_rows WHERE source_file=? AND period_month=? AND report_type='PNLT' "
               "AND dim1 IN (?,?,?,?,?)", (src, period, DIM_CKTM, DIM_TRALAI, DIM_GIAMGIA, DIM_KHAC, DIM_TONG))
    if r is None:
        db.commit()
        return {"ok": False, "skip": True}
    # cong_ty: suy TẤT ĐỊNH từ THƯ MỤC NGUỒN (contract._COMPANY_FOLDER_ALIAS) trước, chỉ fallback
    # sang cong_ty của dòng cùng-khối khi thư mục không nằm trong map.
    # BUG đã bắt 2026-08-05: khối Vận tải Taxi Xanh có 3 PHÁP NHÂN RIÊNG (XVP / HTX_XVP / HTX_XTQ)
    # cùng nộp file đặt tên 'B.6.XVP...', chỉ THƯ MỤC phân biệt được. Lấy `twin` bằng
    # (period, khoi) + LIMIT 1 KHÔNG ORDER BY nên bốc pháp nhân bất kỳ trong khối -> trên prod cả 3
    # thư mục đều bị đóng dấu lẫn XVP/HTX_XTQ và HTX_XVP KHÔNG có dòng nào. Khối-tổng vẫn đúng
    # (revenue._giamtru_of group theo khoi) nhưng view lọc theo Công ty và user bị giới hạn data
    # scope theo pháp nhân thì đọc sai pháp nhân.
    ct = cong_ty
    if not ct:
        from servers.common import contract as C
        from servers.common import source_catalog as SC
        ct = C.resolve_company(raw=SC.raw_company_from_path(path),
                               file_name=os.path.basename(path))
    twin = db.execute("SELECT dataset_id, ngay, cong_ty FROM raw_rows WHERE period_month=? AND khoi=? "
                      "AND dataset_id IS NOT NULL LIMIT 1", (period, khoi)).fetchone()
    if not twin:
        db.commit()
        return {"ok": False, "error": "chưa có dataset của kỳ này"}
    ct = ct or twin["cong_ty"]
    has_other_tong = db.execute(
        "SELECT 1 FROM raw_rows WHERE source_file=? AND period_month=? AND report_type='PNLT' "
        "AND dim1 IN (?,?) LIMIT 1", (src, period, *_EXISTING_TONG_NAMES)).fetchone()
    dims = [(DIM_CKTM, r["ck_tm"]), (DIM_TRALAI, r["tra_lai"]), (DIM_GIAMGIA, r["giam_gia"]),
            (DIM_KHAC, r["khac"])]
    if not has_other_tong:          # tránh đè/đếm đôi với An Taxi (BCQT PT ghi 'Các khoản...' riêng)
        dims.append((DIM_TONG, r["tong"]))
    for i, (dim, val) in enumerate(dims):
        db.execute(
            "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,cost_center,"
            "period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (twin["dataset_id"], "PNLT", 6100000 + i, twin["ngay"], ct, khoi, None, period,
             val, None, dim, None, None,
             json.dumps({"unit": "ty", "nguon": "KQKD mã 02 (derive_kqkd_giamtru)"},
                        ensure_ascii=False), src))
    db.commit()
    return {"ok": True, **{k: round(v, 6) for k, v in r.items()}, "wrote_tong": not has_other_tong}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--khoi", default=None)
    ap.add_argument("--cong-ty", dest="cong_ty", default=None)
    a = ap.parse_args()
    print(json.dumps(extract(a.file, a.period, khoi=a.khoi, cong_ty=a.cong_ty),
                     ensure_ascii=False, default=str))

# -*- coding: utf-8 -*-
"""Deriver: cụm "Phân loại phải thu theo kỳ hạn thu nợ (aging)" — Khối 9, MÀN HÌNH III.1 "Công nợ"
(spec CHÍNH THỨC: /home/itadmin/MoTaChiTiet_Dashboard_TaiChinh.docx).

CHỐT LẠI 2026-08-03 (sau khi đối chiếu docx — bản đầu 2026-08-03 tự dựng bảng "Tuổi nợ phải thu
theo đơn vị" KHÔNG có trong spec, đã bỏ): docx định nghĩa field "aging" như sau —
  Thành phần: < 1 tháng / 1–3 / 3–6 / > 6 tháng
  Lấy trường: Dư PT cuối kỳ chia theo tuổi nợ
  Nguồn: Sổ chi tiết công nợ 131 của đơn vị (ngày phát sinh)
  Công thức: = Ngày báo cáo − Ngày phát sinh; nhóm theo dải
Đây CHÍNH LÀ field `aging` (payload keys `tuoi_no_1t`/`aging_13`/`aging_36`/`aging_6p`) đã có sẵn
trong `debt.py::debt_extras()` — trước giờ luôn 0 vì KHÔNG deriver nào ghi các key này (verify: field
tồn tại ở MỌI dòng PTHU nhưng luôn =0, là placeholder từ template, không phải nguồn riêng công ty
nào). KHÔNG phải "tuổi nợ theo hạn nợ" (trong hạn/đến hạn/quá hạn) — đó là khái niệm KHÁC (dựa Ngày
đến hạn) mà 3 đơn vị này không có (Thời hạn nợ ghi 1825 ngày = 5 năm -> luôn "trong hạn").

Nguồn dữ liệu (per-khách-hàng): `received_reports/<FOLDER>/baocaotuoinophaithu/*.xlsx`, sheet "Báo
cáo tuổi nợ_Mẫu(Tuổi nợ phả...)". Cột cần: "Ngày hóa đơn" (= Ngày phát sinh, cột G) + "Tổng nợ phải
thu" (= dư PT cuối kỳ của dòng đó, cột J) — dò theo TÊN header, không cứng chữ cột.

Tính PER DÒNG (không dùng cột quá-hạn-theo-hạn-nợ N-Q có sẵn trong file — khác trục với "aging"):
  age_days = Ngày báo cáo (CUỐI kỳ, vd 2026-06 -> 2026-06-30) − Ngày hóa đơn
  b1  (< 1 tháng):  age_days < 30
  b13 (1–3 tháng):  30 <= age_days < 90
  b36 (3–6 tháng):  90 <= age_days < 180
  b6p (> 6 tháng):  age_days >= 180
rồi CỘNG "Tổng nợ phải thu" của dòng đó vào đúng dải.

report_type MỚI 'PTHU_TUOINO' (additive, không đụng PTHU/PTRA vàng — cùng pattern
derive_congno_advance.py PTHU_ADV/PTRA_ADV). 1 dòng TỔNG/công ty/kỳ, payload dùng ĐÚNG tên field
`debt.py` đã đọc sẵn (`tuoi_no_1t`/`aging_13`/`aging_36`/`aging_6p`) — debt.py cần sửa để GỘP
report_type này vào cùng `pt` khi tính `aging` (xem debt.py, dòng "aging = ...").

_UNITS: dict folder-nguồn -> (cong_ty, khoi, mode). Chốt 2026-08-03 (bản 2, mở rộng thêm HO/TRẠM SẠC/
GLOBAL AI theo spec mới — CHƯA có file thật, chỉ thêm cấu hình trước) — CHÚ Ý: HO và TRẠM SẠC
CÙNG `cong_ty='TC'` (khác `khoi`) -> PHẢI lưu khoi tường minh ở đây (KHÔNG suy qua "twin" theo
cong_ty như bản đầu, vì twin sẽ vớ nhầm khoi của 1 trong 2 khi cong_ty trùng nhau). DELETE khi ghi
lại cũng scope theo `source_file` (KHÔNG theo cong_ty+period) — cùng lý do, tránh HO đè/xoá nhầm
dòng của TRẠM SẠC. Thêm đơn vị mới: 1 dòng vào dict này (verify cong_ty/khoi qua
`SELECT DISTINCT cong_ty,khoi FROM raw_rows WHERE source_file ILIKE '<FOLDER>::%' LIMIT 5` trên DB
coding trước khi thêm, KHÔNG đoán).

`mode` chọn cách tính, vì 2 nhóm đơn vị có file nguồn KHÁC TRỤC nhau (2026-08-03, thêm SRVF):
  - "age": XVP/HTX_XTQ/HTX_XVP/HO/TRẠM SẠC/GLOBAL AI — "Thời hạn nợ" ghi 1825 ngày (5 năm) nên
    cột quá-hạn-theo-hạn-nợ (K..Q, xem dưới) LUÔN 0/rỗng, không dùng được -> tính age_days từ
    "Ngày hóa đơn" như spec docx (per dòng, xem _find_cols/vòng lặp bucket b1/b13/b36/b6p).
  - "hanno": SRVF — file có "Thời hạn nợ"/"Ngày đến hạn" THẬT (không phải 5 năm), sheet đã tự
    tính sẵn theo khách hàng các cột (dò theo TÊN, xem _find_cols_hanno):
      J "Tổng nợ phải thu" (= K+L+M, verify được ở dòng 0 = dòng tổng của sheet)
      K "Công nợ trong hạn"      L "Công nợ đến hạn"       M "Công nợ quá hạn" (= N+O+P+Q)
      N quá hạn 1-30 ngày · O >30-90 · P >90-180 · Q >180 ngày (sub-header dòng dưới header chính)
    -> KHÔNG tính age_days per dòng (khác hẳn "age") — CỘNG THẲNG giá trị các cột trên qua mọi
    dòng khách hàng. Payload dùng field riêng (tong_no/trong_han/den_han/qua_han/qh_1_30/qh_30_90/
    qh_90_180/qh_180p) — KHÔNG tái dùng key `tuoi_no_1t`/`aging_13`/`aging_36`/`aging_6p` của mode
    "age" vì ý nghĩa khác nhau (age = tuổi TOÀN BỘ dư nợ tính từ hoá đơn; hanno.N..Q chỉ là phần
    QUÁ HẠN, thiếu phần K+L "chưa đến hạn" — gộp lẫn vào cùng field sẽ làm sai lệch donut "Σ dải =
    Tổng dư PT" của debt.py cho các đơn vị mode "age"). debt.py cần đọc thêm field mới này riêng
    khi muốn hiển thị cho SRVF — CHƯA sửa debt.py ở đây (out of scope, agent khác lo phần đọc/UI).

Chạy (dry-run in tổng, KHÔNG ghi):
  .venv/bin/python scripts/derive_congno_tuoino.py <file.xlsx> --period 2026-06
Ghi thật:
  .venv/bin/python scripts/derive_congno_tuoino.py <file.xlsx> --period 2026-06 --write
"""
import argparse
import calendar
import datetime as dt
import json
import os
import re
import sys
import unicodedata

import openpyxl
import psycopg

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers.common import source_catalog as _SC  # noqa: E402

DB_URL = (os.environ.get("DATABASE_URL") or os.environ.get("TC_DATABASE_URL")
          or "postgresql://tc:tc@localhost:5433/tc_dashboard")

# folder nguồn (thư mục received_reports/<FOLDER>) -> (cong_ty, khoi). CHỈ đơn vị đã có nguồn trong
# spec — đơn vị khác không nằm trong dict này -> derive() no-op (skip=True), không ảnh hưởng gì.
# (cong_ty, khoi) verify qua raw_rows thật kỳ 2026-06 (DB coding 5435) trước khi thêm.
_UNITS = {
    "XANHVINHPHUC": ("XVP", "Khối KD Vận tải Taxi Xanh", "age"),
    "HTXXANHTUYENQUANG": ("HTX_XTQ", "Khối KD Vận tải Taxi Xanh", "age"),
    "HTXXANHVINHPHUC": ("HTX_XVP", "Khối KD Vận tải Taxi Xanh", "age"),
    # Chưa có file thật (2026-08-03) — cấu hình sẵn theo spec mới, verify lại cong_ty/khoi khi có
    # file đầu tiên (dry-run trước khi --write).
    "HO": ("TC", "Khối hỗ trợ tập đoàn", "age"),
    "TRAMSAC": ("TC", "Khối KD Trạm sạc Vgreen", "age"),
    "GLOBALAI": ("GA", "Khối KD Công nghệ", "age"),
    # SRVF: verify 2026-08-03 qua raw_rows thật (PTHU/CDPS) trên DB coding — cong_ty='TC',
    # khoi='Khối KD Vinfast - Showroom' (folder SRVF cũng có dữ liệu cong_ty='VFQN' ở nguồn khác,
    # nhưng file baocaotuoino/*.xlsx hiện có TẤT CẢ đều tên "B.1.TC...." -> chỉ company TC).
    "SRVF": ("TC", "Khối KD Vinfast - Showroom", "hanno"),
    # An Taxi/An Khách sạn: verify 2026-08-04 qua raw_rows thật trên DB coding — CÙNG cong_ty='AAG'
    # nhưng khác khoi (giống HO/TRẠM SẠC cùng 'TC') -> khoi lấy tường minh ở đây. File "Báo cáo tuổi
    # nợ" của 2 đơn vị này có "Ngày đến hạn" thật + cột trong-hạn/đến-hạn/quá-hạn (giống SRVF) ->
    # mode 'hanno'. Cột lệch chữ so với SRVF (thêm cột NỢ/CÓ chen giữa J và L) nhưng _find_cols_hanno
    # dò theo TÊN header nên không cần sửa gì thêm, chỉ _find_sheet cần khớp thêm dạng tên sheet
    # "AN T{mm}.26" (xem _find_sheet). ANKHACHSAN CHƯA có file "Báo cáo tuổi nợ" nào (2026-08-04,
    # chỉ có báo cáo tài chính + tài sản cố định) — cấu hình sẵn, chưa verify được cong_ty/khoi qua
    # report_type PTHU_TUOINO (verify qua PTHU/CDPS khác thay, khớp raw_rows thật).
    "ANTAXI": ("AAG", "Khối KD Dịch vụ An Taxi", "hanno"),
    "ANKHACHSAN": ("AAG", "Khối KD Dịch vụ An KS", "hanno"),
}


def _nd(s):
    s = str(s or "").strip().lower().replace("đ", "d")
    return "".join(ch for ch in unicodedata.normalize("NFD", s)
                   if unicodedata.category(ch) != "Mn")


def _num(v):
    return float(v) if isinstance(v, (int, float)) else None


def _source_id(path):
    """folder::basename — DÙNG THẲNG source_catalog.source_id_from_path() (agent_cli._source_id
    cũng gọi hàm này) thay vì tự đoán tên thư mục con. BUG đã gặp (2026-08-03): bản tự viết trước
    loại trừ cứng tên thư mục con 'BAOCAOTUOINOPHAITHU' — HTX_XTQ/HTX_XVP đổi thư mục con thành
    'baocaotuoino' (rút gọn) khi nạp thêm 6 tháng, làm _source_id cũ lấy NHẦM 'baocaotuoino' làm
    folder-đơn-vị thay vì 'HTXXANHTUYENQUANG'/'HTXXANHVINHPHUC'. Hàm thật lấy đúng parts[0] sau
    RECEIVED_DIR (không quan tâm tên/số cấp thư mục con), không lặp lại lỗi này."""
    folder = (_SC.raw_company_from_path(path) or "").upper()
    return f"{folder}::{os.path.basename(path)}"


def _find_sheet(wb, period=None):
    """Tên sheet KHÔNG cố định — file kỳ 06/2026 đặt tên đầy đủ 'Báo cáo tuổi nợ_Mẫu(...)', các
    kỳ trước (T01-T05, phát hiện 2026-08-03 khi HTX_XTQ/HTX_XVP có thêm 6 tháng) chỉ đặt 'T1'..'T6'
    (số trùng THÁNG); An Taxi (phát hiện 2026-08-04) đặt kiểu 'AN T6.26' (có tiền tố + hậu tố năm).
    Ưu tiên tên mô tả; fallback khớp token 't{mm}' ở BẤT KỲ ĐÂU trong tên (biên không phải chữ số,
    tránh 'T1' khớp nhầm 'T10'/'T11') thay vì đòi tên sheet == 'T{mm}' tuyệt đối — không lấy 'T1'
    đầu tiên bừa (cùng bẫy đã gặp ở derive_tscd_hetkhauhao._sheet_thang với sheet 'Tháng N')."""
    sn = next((s for s in wb.sheetnames if "tuoi no" in _nd(s) or "tuoi_no" in _nd(s)), None)
    if sn or not period:
        return sn
    mm = int(period[5:7])
    pat = re.compile(rf"(?<![0-9])t0?{mm}(?![0-9])")
    return next((s for s in wb.sheetnames if pat.search(_nd(s))), None)


def _period_end(period):
    """'YYYY-MM' -> ngày cuối tháng (Ngày báo cáo)."""
    y, m = int(period[:4]), int(period[5:7])
    return dt.date(y, m, calendar.monthrange(y, m)[1])


def _find_cols(rows):
    """Dò header theo TÊN — chỉ cần 'Ngày hóa đơn' (ngày phát sinh) + 'Tổng nợ phải thu' (dư PT).
    Trả (data_start, ngay_i, tong_i) hoặc (None, None, None)."""
    hdr_i = next((i for i, r in enumerate(rows[:6])
                  if any(_nd(c).startswith("tong no phai thu") for c in r if c)), None)
    if hdr_i is None:
        return None, None, None
    h = rows[hdr_i]
    ngay_i = next((j for j, c in enumerate(h) if c and _nd(c).startswith("ngay hoa don")), None)
    tong_i = next((j for j, c in enumerate(h) if c and _nd(c).startswith("tong no phai thu")), None)
    ma_i = next((j for j, c in enumerate(h) if c and _nd(c) == "ma khach"), None)
    if ngay_i is None or tong_i is None:
        return None, None, None
    return hdr_i + 2, ngay_i, tong_i, (ma_i if ma_i is not None else 4)


def _find_cols_hanno(rows):
    """Dò cột cho mode 'hanno' (SRVF) — header CHÍNH ở hàng chứa 'Tổng nợ phải thu' (giống
    _find_cols), nhưng 4 dải quá hạn theo ngày nằm ở SUB-HEADER hàng NGAY DƯỚI (cell 'Công nợ quá
    hạn' bị merge ngang N:Q ở hàng chính, tên dải thật '1-30 ngày'/'>30-90'/'>90-180'/'>180 ngày'
    chỉ xuất hiện ở hàng phụ) — verify bằng file thật SRVF/baocaotuoino T06/2026.
    Trả dict cột hoặc None nếu không đủ cột (fallback sang mode 'age' ở derive())."""
    hdr_i = next((i for i, r in enumerate(rows[:6])
                  if any(_nd(c).startswith("tong no phai thu") for c in r if c)), None)
    if hdr_i is None or hdr_i + 1 >= len(rows):
        return None
    h, sub = rows[hdr_i], rows[hdr_i + 1]

    def _idx(row, pred):
        return next((j for j, c in enumerate(row) if c and pred(_nd(c))), None)

    cols = {
        "tong": _idx(h, lambda s: s.startswith("tong no phai thu")),
        "trong_han": _idx(h, lambda s: s.startswith("cong no trong han")),
        "den_han": _idx(h, lambda s: s.startswith("cong no den han")),
        "qua_han": _idx(h, lambda s: s.startswith("cong no qua han")),
        "qh_1_30": _idx(sub, lambda s: s.startswith("1-30")),
        "qh_30_90": _idx(sub, lambda s: "30-90" in s),
        "qh_90_180": _idx(sub, lambda s: "90-180" in s),
        "qh_180p": _idx(sub, lambda s: s.startswith("180") or s.startswith(">180")),
    }
    if any(v is None for v in cols.values()):
        return None
    ma_i = _idx(h, lambda s: s == "ma khach")
    cols["ma"] = ma_i if ma_i is not None else 4
    cols["data_start"] = hdr_i + 2
    return cols


def _twin_attrs(cur, period):
    """dataset_id/ngay từ 1 dòng BẤT KỲ đã nạp trong CÙNG period_month (dataset/ngay là cấp ĐỘ KỲ,
    dùng chung mọi công ty trong tháng đó — KHÔNG cần lọc theo cong_ty/khoi). `khoi` LẤY THẲNG từ
    `_UNITS` (không suy qua twin): nhiều đơn vị CÙNG cong_ty khác khoi (vd HO/TRẠM SẠC đều 'TC') ->
    twin theo cong_ty sẽ vớ nhầm khoi của đơn vị khác cùng mã công ty.
    `ngay` BẮT BUỘC không None: report_type nằm trong `_SNAP_RT` (metrics/_shared.py) nên `_rows()`
    chọn snapshot theo MAX(ngay); ngay=None -> _rows() trả [] câm (không lỗi)."""
    cur.execute(
        "SELECT dataset_id, ngay FROM raw_rows WHERE period_month=%s "
        "AND ngay IS NOT NULL ORDER BY (report_type='PTHU') DESC LIMIT 1",
        (period,))
    return cur.fetchone()


def _agg_age(rows, report_date):
    """mode 'age' (XVP/HTX/HO/TRẠM SẠC/GLOBAL AI) — bucket theo tuổi hoá đơn, xem docstring đầu file."""
    data_start, ngay_i, tong_i, ma_i = _find_cols(rows)
    if data_start is None:
        return None
    agg = {"b1": 0.0, "b13": 0.0, "b36": 0.0, "b6p": 0.0}
    n_rows = 0
    for r in rows[data_start:]:
        if not r or ma_i >= len(r) or not r[ma_i]:
            continue
        ngay = r[ngay_i] if ngay_i < len(r) else None
        tong = _num(r[tong_i]) if tong_i < len(r) else None
        if not isinstance(ngay, (dt.date, dt.datetime)) or tong is None:
            continue
        n_rows += 1
        age_days = (report_date - (ngay.date() if isinstance(ngay, dt.datetime) else ngay)).days
        bucket = "b1" if age_days < 30 else "b13" if age_days < 90 else "b36" if age_days < 180 else "b6p"
        agg[bucket] += tong
    tong_all = round(sum(agg.values()) * 1e-9, 9)
    payload = {"tuoi_no_1t": round(agg["b1"] * 1e-9, 9), "aging_13": round(agg["b13"] * 1e-9, 9),
               "aging_36": round(agg["b36"] * 1e-9, 9), "aging_6p": round(agg["b6p"] * 1e-9, 9),
               "unit": "ty"}
    return {"n_rows": n_rows, "tong_all": tong_all,
            "agg_ty": {k: round(v * 1e-9, 9) for k, v in agg.items()}, "payload": payload}


def _agg_hanno(rows):
    """mode 'hanno' (SRVF) — CỘNG THẲNG các cột trong-hạn/đến-hạn/quá-hạn (J..Q) đã có sẵn theo
    khách hàng trong sheet, KHÔNG tính age_days (khác trục 'age', xem docstring đầu file)."""
    cols = _find_cols_hanno(rows)
    if cols is None:
        return None
    keys = ("tong", "trong_han", "den_han", "qua_han", "qh_1_30", "qh_30_90", "qh_90_180", "qh_180p")
    agg = {k: 0.0 for k in keys}
    n_rows = 0
    for r in rows[cols["data_start"]:]:
        if not r or cols["ma"] >= len(r) or not r[cols["ma"]]:
            continue
        vals = {k: _num(r[cols[k]]) if cols[k] < len(r) else None for k in keys}
        if vals["tong"] is None:
            continue
        n_rows += 1
        for k in keys:
            agg[k] += vals[k] or 0.0
    tong_all = round(agg["tong"] * 1e-9, 9)
    payload = {"tong_no": round(agg["tong"] * 1e-9, 9), "trong_han": round(agg["trong_han"] * 1e-9, 9),
               "den_han": round(agg["den_han"] * 1e-9, 9), "qua_han": round(agg["qua_han"] * 1e-9, 9),
               "qh_1_30": round(agg["qh_1_30"] * 1e-9, 9), "qh_30_90": round(agg["qh_30_90"] * 1e-9, 9),
               "qh_90_180": round(agg["qh_90_180"] * 1e-9, 9), "qh_180p": round(agg["qh_180p"] * 1e-9, 9),
               "unit": "ty"}
    return {"n_rows": n_rows, "tong_all": tong_all, "agg_ty": {k: round(v * 1e-9, 9) for k, v in agg.items()},
            "payload": payload}


def derive(path, period, write=False):
    folder = _source_id(path).split("::", 1)[0]
    unit = _UNITS.get(folder)
    if not unit:
        return {"ok": False, "skip": True}
    cong_ty, khoi, mode = unit
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sn = _find_sheet(wb, period)
        if not sn:
            return {"ok": False, "error": "không thấy sheet 'Báo cáo tuổi nợ'"}
        rows = [list(r) for r in wb[sn].iter_rows(values_only=True)]
    finally:
        wb.close()

    result = _agg_hanno(rows) if mode == "hanno" else _agg_age(rows, _period_end(period))
    if result is None:
        err = ("không dò được cột trong-hạn/đến-hạn/quá-hạn (J..Q)" if mode == "hanno"
               else "không dò được cột 'Ngày hóa đơn'/'Tổng nợ phải thu'")
        return {"ok": False, "error": err}
    n_rows, tong_all, payload = result["n_rows"], result["tong_all"], result["payload"]

    out = {"file": os.path.basename(path), "period": period, "cong_ty": cong_ty, "mode": mode,
           "n_rows": n_rows, "tong_ty": tong_all, "agg_ty": result["agg_ty"]}

    if write:
        source_file = _source_id(path)
        conn = psycopg.connect(DB_URL)
        try:
            cur = conn.cursor()
            attrs = _twin_attrs(cur, period)
            if not attrs:
                out["error"] = f"không thấy dòng nào có ngay (period={period}) để suy dataset_id/ngay -> BỎ ghi"
                return out
            dataset_id, ngay = attrs
            # idempotent: xoá bản cũ CÙNG source_file (KHÔNG cong_ty+period — vài đơn vị dùng CHUNG
            # cong_ty như HO/TRẠM SẠC đều 'TC', xoá theo cong_ty sẽ xoá NHẦM sang đơn vị khác cùng mã)
            cur.execute("DELETE FROM raw_rows WHERE report_type='PTHU_TUOINO' AND source_file=%s",
                        (source_file,))
            cur.execute(
                "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
                "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, source_file) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (dataset_id, "PTHU_TUOINO", 6100000, ngay, cong_ty, khoi, None, period,
                 tong_all, None, cong_ty, None, None,
                 json.dumps(payload, ensure_ascii=False), source_file))
            conn.commit()
            out["written"] = 1
        finally:
            conn.close()
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--write", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    print(json.dumps(derive(a.file, a.period, a.write), ensure_ascii=False, indent=2))

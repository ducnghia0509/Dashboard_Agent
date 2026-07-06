# -*- coding: utf-8 -*-
"""Bóc SỐ DƯ TIỀN từ Báo cáo Tiền tập đoàn (sheet '..._SD TIỀN') -> template 03B_SODU_TIEN -> SDT.

Nguồn (Nguồn dữ liệu lên Dashboard.xlsx, dòng DÒNG TIỀN): 'Baocaothuchi / 9.HT_NGANHANG'.
Layout SD TIỀN: phần LOẠI TIỀN (I=TIỀN VAY, II=TIỀN GỬI, III=TIỀN MẶT, IV=BẢO LÃNH, V=LC) ->
mỗi phần có dòng CÔNG TY (col2, col ngân hàng rỗng), cột 'ĐẾN NGÀY HIỆN TẠI' (index 5) = số dư.
=> pivot phần->cột theo từng pháp nhân. Chạy:
   .venv/bin/python scripts/extract_sodu_tien.py <file.xlsx> --period 2026-01
"""
import argparse
import os
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers import template_filler as tf  # noqa: E402

# Tên công ty trong SD TIỀN -> mã pháp nhân (MD_CONGTY). Dòng tiền/số dư gán THEO PHÁP NHÂN.
_CO = {"thịnh cường": "TC", "hưng thịnh": "HT", "vinfast quảng ninh": "VFQN",
       "xanh vĩnh phúc": "XVP", "an taxi": "AAG"}
_SEC = {"TIỀN VAY": "vay", "TIỀN GỬI": "gui", "TIỀN MẶT": "mat", "BẢO LÃNH": "bl", "LC": "lc"}
_SD_COL = 5  # 'ĐẾN NGÀY HIỆN TẠI' (số dư cuối)


def _sd_sheet(wb):
    for ws in wb.worksheets:
        if "SD TI" in ws.title.upper() or "SỐ DƯ TI" in ws.title.upper():
            return ws
    return None


def extract(path: str, period: str, cong_ty: str = None) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _sd_sheet(wb)
    if ws is None:
        return {"ok": False, "error": "Không thấy sheet SD TIỀN"}
    data, sec = {}, None
    for r in ws.iter_rows(values_only=True):
        c1 = str(r[1]).strip() if len(r) > 1 and r[1] not in (None, "") else ""
        c2 = str(r[2]).strip() if len(r) > 2 and r[2] not in (None, "") else ""
        c3 = r[3] if len(r) > 3 else None
        c5 = r[_SD_COL] if len(r) > _SD_COL else None
        if c1 in _SEC:
            sec = _SEC[c1]
            continue
        # dòng CÔNG TY tổng: col2 = tên cty, col1 rỗng (không phải mã phụ), col ngân hàng rỗng
        if sec and c2 and not c1 and (c3 in (None, "")):
            code = _CO.get(c2.lower())
            if code and isinstance(c5, (int, float)):
                data.setdefault(code, {})
                data[code][sec] = data[code].get(sec, 0.0) + c5
    if not data:
        return {"ok": False, "error": "Không bóc được số dư công ty nào"}
    recs = []
    for code, d in data.items():
        recs.append({"Kỳ": period, "Đơn vị": code,
                     "Tiền mặt (tỷ)": round(d.get("mat", 0) / 1e9, 6),
                     "Tiền gửi NH (tỷ)": round(d.get("gui", 0) / 1e9, 6),
                     "Ngoại bảng: LC (tỷ)": round(d.get("lc", 0) / 1e9, 6),
                     "Bảo lãnh thanh toán (tỷ)": round(d.get("bl", 0) / 1e9, 6),
                     "Số dư tiền vay (tỷ) — đối chiếu 04_VAY": round(d.get("vay", 0) / 1e9, 6)})
    out = os.path.join(tf.FILLED_DIR, f"THUCHI_{period}_03B_SODU_TIEN.xlsx")
    tf.fill("03B_SODU_TIEN", recs, out)
    imp = tf.import_filled(out, cong_ty=cong_ty)
    return {"ok": True, "companies": list(data), "rows": imp.get("rows_imported"),
            "by_type": imp.get("by_type"), "out": out}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--cong-ty", dest="cong_ty", default=None)
    a = ap.parse_args()
    import json
    print(json.dumps(extract(a.file, a.period, a.cong_ty), ensure_ascii=False))

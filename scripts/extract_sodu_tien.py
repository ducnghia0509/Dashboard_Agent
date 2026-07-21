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
from servers.common import source_catalog as SC  # noqa: E402
from servers.common import org  # noqa: E402  — danh mục tổ chức (nguồn define duy nhất)

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
    data, sec, seen = {}, None, set()
    _dachi = None   # 'Đã chi chưa có chứng từ' (dưới mục TIỀN MẶT) — cảnh báo, gán cho TC (guide 21/7)
    from servers.common import be_bridge as bb   # remove_diacritics: 'đã…chứng từ' -> 'da…chung tu'
    _norm = lambda s: bb.remove_diacritics(str(s or "")).strip().lower()  # noqa: E731
    for r in ws.iter_rows(values_only=True):
        c1 = str(r[1]).strip() if len(r) > 1 and r[1] not in (None, "") else ""
        c2 = str(r[2]).strip() if len(r) > 2 and r[2] not in (None, "") else ""
        c3 = r[3] if len(r) > 3 else None
        c4 = r[4] if len(r) > 4 else None    # ĐẦU KỲ
        c5 = r[_SD_COL] if len(r) > _SD_COL else None
        if c1 in _SEC:
            # Nguồn liệt kê 'TIỀN VAY' 2 lần (khối 'đầu kỳ 30/04' + khối 'đầu kỳ 01/04', CÙNG cột
            # 'đến ngày hiện tại') -> nếu cộng cả 2 thì số dư vay bị NHÂN ĐÔI. CHỈ lấy khối ĐẦU của
            # mỗi loại tiền; section lặp -> sec=None (bỏ). Tiền gửi/mặt chỉ 1 khối nên không đổi.
            s = _SEC[c1]
            sec = None if s in seen else s
            seen.add(s)
            continue
        # 'Đã chi nhưng chưa có chứng từ' (memo dưới mục TIỀN MẶT) -> cảnh báo Thịnh Cường (TC). Bắt
        # theo TÊN dòng (KHÔNG phải mã công ty). Lấy cả cuối kỳ (col5) & đầu kỳ (col4). Chỉ khối 'mat'.
        if sec == "mat" and "da chi" in _norm(c2) and "chung tu" in _norm(c2):
            _dachi = {"cuoi": c5 if isinstance(c5, (int, float)) else 0.0,
                      "dau": c4 if isinstance(c4, (int, float)) else 0.0, "ten": c2}
            continue
        # dòng CÔNG TY tổng: col2 = tên cty, col1 rỗng (không phải mã phụ), col ngân hàng rỗng
        if sec and c2 and not c1 and (c3 in (None, "")):
            code = org.company_code(c2)
            if code and isinstance(c5, (int, float)):
                data.setdefault(code, {})
                data[code][sec] = data[code].get(sec, 0.0) + c5
    if not data:
        return {"ok": False, "error": "Không bóc được số dư công ty nào"}
    recs = []
    for code, d in data.items():
        recs.append({"Kỳ": period, "Đơn vị": code,
                     "Tiền mặt (tỷ)": round(d.get("mat", 0) / 1e9, 9),
                     "Tiền gửi NH (tỷ)": round(d.get("gui", 0) / 1e9, 9),
                     "Ngoại bảng: LC (tỷ)": round(d.get("lc", 0) / 1e9, 9),
                     "Bảo lãnh thanh toán (tỷ)": round(d.get("bl", 0) / 1e9, 9),
                     "Số dư tiền vay (tỷ) — đối chiếu 04_VAY": round(d.get("vay", 0) / 1e9, 9)})
    out = os.path.join(tf.FILLED_DIR, f"THUCHI_{period}_03B_SODU_TIEN.xlsx")
    tf.fill("03B_SODU_TIEN", recs, out)
    imp = tf.import_filled(out, cong_ty=cong_ty, source_file=SC.source_id_from_path(path))
    # 'Đã chi nhưng chưa có chứng từ' -> report_type ADDITIVE 'DACHI_CCT' cho TC (Thịnh Cường), chỉ số
    # cảnh báo màn Dòng tiền (guide 21/7). Direct-insert (không qua template — tránh sửa Template_chuan):
    # amount=cuối kỳ, amount2=đầu kỳ. Idempotent theo (source_file, period). Twin = 1 dòng SDT TC vừa nạp.
    src = SC.source_id_from_path(path)
    if _dachi is not None:
        from servers.common import be_bridge as bb
        import json as _json
        _db = bb.db.get_db()
        _db.execute("DELETE FROM raw_rows WHERE source_file=? AND period_month=? AND report_type=?",
                    (src, period, "DACHI_CCT"))
        _tw = _db.execute("SELECT dataset_id, ngay, khoi FROM raw_rows WHERE source_file=? AND "
                          "period_month=? AND report_type='SDT' AND cong_ty='TC' LIMIT 1",
                          (src, period)).fetchone()
        if _tw:
            _db.execute(
                "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,cost_center,"
                "period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (_tw["dataset_id"], "DACHI_CCT", 6000000, _tw["ngay"], "TC", _tw["khoi"], None, period,
                 round((_dachi["cuoi"] or 0) * 1e-9, 9), round((_dachi["dau"] or 0) * 1e-9, 9),
                 "Đã chi nhưng chưa có chứng từ", None, None,
                 _json.dumps({"unit": "ty", "nguon": "TC01_SD TIỀN - Đã chi chưa có chứng từ"},
                             ensure_ascii=False), src))
            _db.commit()
    return {"ok": True, "companies": list(data), "rows": imp.get("rows_imported"),
            "by_type": imp.get("by_type"), "out": out, "da_chi_cct": _dachi}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--cong-ty", dest="cong_ty", default=None)
    a = ap.parse_args()
    import json
    print(json.dumps(extract(a.file, a.period, a.cong_ty), ensure_ascii=False))

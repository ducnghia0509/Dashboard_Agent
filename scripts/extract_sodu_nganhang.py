# -*- coding: utf-8 -*-
"""SỐ DƯ NGÂN HÀNG theo NGÀY (khả dụng / phong tỏa) — 2 chart mới màn "Dòng tiền"
(mapping dongtienrealtime.xlsx dòng 54-70, report_type mới "SDNH").

Nguồn: Connect_VPS/received_reports/THUCHI/sodu<entity>/<entity>_<BANK>.(xlsx|xls) — file GHI ĐÈ
mỗi lần kế toán cập nhật (KHÔNG giữ lịch sử ngày cũ), 1 file/ngân hàng/pháp nhân. Bốn pháp nhân có
dữ liệu: Hưng Thịnh / Thịnh Cường / Xanh Vĩnh Phúc / VFQN ("An" chưa có nguồn — mapping ghi "chưa có").

Bố cục theo NGÂN HÀNG (không phụ thuộc pháp nhân), dò theo TEXT header/label (không hard-code cột)
— đã xác minh trên toàn bộ 15 file hiện có:
  · BIDV / VPBank (bảng nhiều dòng, 1 dòng/tài khoản): cột "Số dư khả dụng" luôn có; cột phong tỏa
    là "Số tiền phong tỏa" (BIDV) HOẶC "Số tiền khoanh giữ" (VPBank) — cộng dồn cả 2 cột qua mọi
    dòng tài khoản.
  · MB (bảng nhiều dòng): chỉ có cột "Số dư tài khoản" — dùng làm khả dụng, KHÔNG có khái niệm
    phong tỏa (mapping ghi "- Không có").
  · VCB / Bắc Á (key-value 1 tài khoản): 2 dòng nhãn "Số dư khả dụng" / "Số tiền phong tỏa"
    (hoặc "khoanh giữ"), giá trị ở ô liền sau.

Ghi TRỰC TIẾP raw_rows (KHÔNG qua golden template) — theo đúng tiền lệ DACHI_CCT trong
extract_tien.py: đây là 1 chỉ tiêu mới đơn giản, không đáng thêm sheet vào Template_chuan.xlsx dùng
chung. `dataset_id` MƯỢN từ dòng SDT mới nhất cùng (cong_ty, period_month) — cùng kỹ thuật DACHI_CCT.
Idempotent theo (report_type='SDNH', ngay, cong_ty, dim1): chạy nhiều lần/ngày chỉ giữ lần cuối.

Chạy: .venv/bin/python scripts/extract_sodu_nganhang.py [--commit] [--ngay yyyy-mm-dd]
"""
import argparse
import datetime
import glob
import os
import sys

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(_HERE))
from servers.common import be_bridge as bb  # noqa: E402

_WS = os.path.normpath(os.path.join(_HERE, "..", ".."))
SCAN_ROOT = os.path.join(_WS, "Connect_VPS", "received_reports", "THUCHI")

_FOLDER_CTY = {"soduhungthinh": "HT", "soduthinhcuong": "TC",
               "soduxanhvinhphuc": "XVP", "soduvfqn": "VFQN"}
_BANK_DISPLAY = {"BIDV": "BIDV", "MB": "MB", "VPBANK": "VP", "VCB": "VCB", "BAC A": "Bắc Á"}


def _norm(s):
    return bb.remove_diacritics("" if s is None else str(s)).strip().lower()


def _num(v):
    """Chấp nhận số thật lẫn chuỗi có dấu phẩy ('224,043,264') hoặc '0'."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace(" ", "")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _find_col(row, *kws):
    for j, c in enumerate(row):
        if c is None:
            continue
        s = _norm(c)
        if s and all(k in s for k in kws):
            return j
    return None


def _is_total_row(row):
    return any(("tong cong" in _norm(c) or "tong so" in _norm(c)) for c in row)


def _load_rows(path):
    """-> list[list[value]], engine-agnostic (.xls qua xlrd, .xlsx/.xlsm qua openpyxl)."""
    if path.lower().endswith(".xls"):
        import xlrd
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        return [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def read_balance(rows):
    """-> (kha_dung, phong_toa) tỷ VND thô (chưa /1e9). phong_toa=None nếu NH không có khái niệm này."""
    # 1) BẢNG nhiều dòng: cột "Số dư khả dụng" (BIDV/VPBank) hoặc "Số dư tài khoản" (MB).
    for i, r in enumerate(rows[:15]):
        col = _find_col(r, "so du kha dung")
        found_kd_col = col is not None
        if col is None:
            col = _find_col(r, "so du tai khoan")
        if col is None:
            continue
        c_pt = _find_col(r, "so tien phong toa") or _find_col(r, "so tien khoanh giu") \
            if found_kd_col else None
        kd = pt = 0.0
        any_row = False
        for r2 in rows[i + 1:]:
            if _is_total_row(r2):
                continue
            v = _num(r2[col]) if col < len(r2) else None
            if v is None:
                continue
            any_row = True
            kd += v
            if c_pt is not None:
                vp = _num(r2[c_pt]) if c_pt < len(r2) else None
                if vp is not None:
                    pt += vp
        if any_row:
            return kd, (pt if c_pt is not None else None)
    # 2) KEY-VALUE 1 tài khoản (VCB/Bắc Á): nhãn ở 1 ô, giá trị ở ô liền sau CÙNG dòng.
    kd = pt = None
    for r in rows:
        for j, c in enumerate(r):
            if not isinstance(c, str):
                continue
            s = _norm(c)
            if kd is None and "so du kha dung" in s and j + 1 < len(r):
                kd = _num(r[j + 1])
            if pt is None and ("so tien phong toa" in s or "so tien khoanh giu" in s) and j + 1 < len(r):
                pt = _num(r[j + 1])
    return kd, pt


def bank_of_filename(stem):
    """'<entity>_<BANK...>' -> mã ngân hàng chuẩn hoá (BIDV/MB/VPBANK/VCB/BAC A) hoặc None."""
    if "_" not in stem:
        return None
    tok = _norm(stem.split("_", 1)[1])
    if "vpbank" in tok or tok == "vp":
        return "VPBANK"
    if "bidv" in tok:
        return "BIDV"
    if tok == "mb" or "mb" in tok.split():
        return "MB"
    if "vcb" in tok or "vietcombank" in tok:
        return "VCB"
    if "bac a" in tok or "bacabank" in tok:
        return "BAC A"
    return None


def files_can_doc():
    """[(cong_ty, bank_display, path), ...] — 1 file/ngân hàng/pháp nhân, quét sodu<entity>/."""
    out = []
    for folder, cty in _FOLDER_CTY.items():
        for f in sorted(glob.glob(os.path.join(SCAN_ROOT, folder, "*"))):
            ten = os.path.basename(f)
            if not ten.lower().endswith((".xlsx", ".xlsm", ".xls")) or ten.startswith("~$"):
                continue
            stem = os.path.splitext(ten)[0]
            bank = bank_of_filename(stem)
            if bank is None:
                continue
            out.append((cty, _BANK_DISPLAY[bank], f))
    return out


def _dataset_id_for(db, cty, period):
    row = db.execute(
        "SELECT dataset_id FROM raw_rows WHERE report_type='SDT' AND cong_ty=? AND period_month=? "
        "ORDER BY ngay DESC LIMIT 1", (cty, period)).fetchone()
    return row["dataset_id"] if row else None


def run(commit, ngay=None):
    ngay = ngay or datetime.date.today().isoformat()
    period = ngay[:7]
    db = bb.db.get_db()
    results, skipped = [], []
    for cty, bank, f in files_can_doc():
        try:
            rows = _load_rows(f)
        except Exception as exc:  # noqa: BLE001 — file lỗi 1 ngân hàng không chặn các ngân hàng khác
            skipped.append({"cong_ty": cty, "bank": bank, "file": f, "error": str(exc)})
            continue
        kd, pt = read_balance(rows)
        if kd is None:
            skipped.append({"cong_ty": cty, "bank": bank, "file": f, "error": "khong_doc_duoc_so_du"})
            continue
        ds = _dataset_id_for(db, cty, period)
        if ds is None:
            skipped.append({"cong_ty": cty, "bank": bank, "file": f,
                             "error": f"chua_co_dataset_SDT_ky_{period}"})
            continue
        results.append({"cong_ty": cty, "bank": bank, "dataset_id": ds,
                         "kha_dung_ty": round(kd / 1e9, 9),
                         "phong_toa_ty": round(pt / 1e9, 9) if pt is not None else None})
        if commit:
            db.execute("DELETE FROM raw_rows WHERE report_type='SDNH' AND ngay=? AND cong_ty=? "
                       "AND dim1=?", (ngay, cty, bank))
            db.execute(
                "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,cost_center,"
                "period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (ds, "SDNH", 7000000, ngay, cty, None, None, period,
                 round(kd / 1e9, 9), round(pt / 1e9, 9) if pt is not None else None,
                 bank, None, None, None, f))
    if commit:
        db.commit()
    return {"ngay": ngay, "ok": results, "skipped": skipped}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--ngay", default=None, help="yyyy-mm-dd (mặc định hôm nay)")
    a = ap.parse_args()
    import json
    out = run(a.commit, a.ngay)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    print(f"\n{'ĐÃ GHI' if a.commit else 'DRY-RUN'}: {len(out['ok'])} (đơn vị,ngân hàng), "
          f"{len(out['skipped'])} bỏ qua. {'' if a.commit else 'Chạy --commit để ghi.'}")


if __name__ == "__main__":
    main()

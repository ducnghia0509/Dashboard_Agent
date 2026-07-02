# -*- coding: utf-8 -*-
"""Sinh kpi_glossary.json tu ../DashBoard_AI/guideline.xlsx (sheet 'Quan tri Tai chinh').

Cau truc sheet that (xac nhan bang cach doc tung o rieng le, khong theo dong phang):
  Cot A = nhom bao cao lon (vd 'HIEU QUA KINH DOANH', 'DONG TIEN', 'VAY VA LAI VAY'),
          merged qua nhieu dong -> phai forward-fill.
  Cot B = nhom con (vd 'DOANH THU', 'GIA VON - LOI NHUAN GOP'), cung merged -> forward-fill.
  Cot C = ten chi tieu phan tich, doi khi merged qua vai dong (nhieu 'chieu phan tich'
          cung 1 chi tieu) -> forward-fill.
  Cot D..J = Chieu phan tich / Dac trung / Cong thuc / Don vi / Nguon du lieu /
             Chi tieu canh bao do / Hien bang tong quan - gia tri rieng cho tung dong,
             KHONG forward-fill.
Moi dong co it nhat 1 trong D..J la 1 "chi tiet" (detail) cua 1 chi tieu.

Chay: python scripts/gen_kpi_glossary.py
Ghi de DashBoard_Agent/kpi_glossary.json.
"""
import json
import os
import sys

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
GUIDELINE_XLSX = os.path.normpath(os.path.join(ROOT, "..", "DashBoard_AI", "guideline.xlsx"))
OUT_PATH = os.path.join(ROOT, "kpi_glossary.json")
SHEET_NAME = "Quản trị Tài Chính"

FIRST_DATA_ROW = 5  # dòng 1-4 = tiêu đề + header cột

FOLLOWUP_MARKERS = [
    "cần chốt", "chưa cập nhật", "chưa có", "check lại", "hỏi lại",
    "cần trao đổi", "note:", "cần hỏi",
]
NEW_SOURCE_MARKERS = [
    "tk131", "tk331", "tscđ", "cđsps", "ngoại bảng", "aging",
    "tuổi nợ", "luân chuyển", "khấu hao", "hao mòn", "ngân hàng",
    "kỳ hạn", "đến hạn",
]


def _s(v):
    return "" if v is None else str(v).strip()


def _needs_followup(rec) -> bool:
    hay = " ".join([
        rec["chi_tieu"], rec["dac_trung"], rec["cong_thuc"], rec["nguon_du_lieu"],
    ]).lower()
    return any(m in hay for m in FOLLOWUP_MARKERS + NEW_SOURCE_MARKERS)


def build_glossary(path=GUIDELINE_XLSX, sheet=SHEET_NAME):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]

    records = []
    group, subgroup, chi_tieu = "", "", ""
    for row in range(FIRST_DATA_ROW, ws.max_row + 1):
        a = _s(ws.cell(row, 1).value)
        b = _s(ws.cell(row, 2).value)
        c = _s(ws.cell(row, 3).value)
        d = _s(ws.cell(row, 4).value)
        e = _s(ws.cell(row, 5).value)
        f = _s(ws.cell(row, 6).value)
        g = _s(ws.cell(row, 7).value)
        h = _s(ws.cell(row, 8).value)
        i = _s(ws.cell(row, 9).value)
        j = _s(ws.cell(row, 10).value)

        if a:
            group = a
        if b:
            subgroup = b
        if c:
            chi_tieu = c

        if not any([d, e, f, g, h, i, j]):
            continue  # dòng chỉ mang tên nhóm/chỉ tiêu, không có chi tiết riêng

        rec = {
            "row": row,
            "nhom_bao_cao": group,
            "nhom_con": subgroup,
            "chi_tieu": chi_tieu,
            "chieu_phan_tich": d,
            "dac_trung": e,
            "cong_thuc": f,
            "don_vi": g,
            "nguon_du_lieu": h,
            "canh_bao_do": i,
            "hien_tong_quan": j,
        }
        rec["needs_followup"] = _needs_followup(rec)
        rec["wired"] = not rec["needs_followup"]
        records.append(rec)

    for idx, rec in enumerate(records, start=1):
        rec["id"] = idx
    return records


def main():
    if not os.path.exists(GUIDELINE_XLSX):
        print(f"Không tìm thấy {GUIDELINE_XLSX}", file=sys.stderr)
        sys.exit(1)
    records = build_glossary()
    with open(OUT_PATH, "w", encoding="utf-8") as fh:
        json.dump(records, fh, ensure_ascii=False, indent=2)
    n_followup = sum(1 for r in records if r["needs_followup"])
    print(f"Đã ghi {OUT_PATH}: {len(records)} record ({n_followup} cần bổ sung nguồn dữ liệu).")


if __name__ == "__main__":
    main()

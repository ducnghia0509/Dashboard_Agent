# -*- coding: utf-8 -*-
"""Sinh kpi_glossary.json TỪ template vàng Template_chuan.xlsx, sheet '00_50CHITIEU'.

NGUỒN SỰ THẬT DUY NHẤT (2026-07): guideline rời (guildline.xlsx) đã ARCHIVE. Bộ 50 chỉ
tiêu chuẩn giờ sống trong Template_chuan!00_50CHITIEU — CÙNG file analyst đọc live qua
contract.guide()/template_contract_info. Regenerate glossary từ đây => QA (glossary_lookup)
và ingest (analyst) không bao giờ lệch nhau.

Sheet 00_50CHITIEU (dòng 1 = tiêu đề, dòng 2 = header cột, dữ liệu từ dòng 3):
  cột: STT | Nhóm chỉ tiêu | Chỉ tiêu (tên đúng) | Giá trị mẫu | Trạng thái đèn |
       Công thức/Logic | ĐVT | Nguồn dữ liệu | Sheet nhập liệu | Cột nguồn |
       Chiều phân tích | Ngưỡng VÀNG | Ngưỡng ĐỎ | Chiều đánh giá
  Dòng tiêu đề nhóm ('▶ 1. Doanh thu (3)') ở cột đầu, cột 'Chỉ tiêu' rỗng -> BỎ.

Chạy: python scripts/gen_kpi_glossary.py   -> ghi đè kpi_glossary.json
"""
import json
import os
import sys
import unicodedata

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
# Ưu tiên env GOLDEN_TEMPLATE (khớp contract.py), mặc định ~/template_trust/Template_chuan.xlsx
TEMPLATE_XLSX = os.environ.get("GOLDEN_TEMPLATE") or os.path.normpath(
    os.path.join(ROOT, "..", "template_trust", "Template_chuan.xlsx"))
OUT_PATH = os.path.join(ROOT, "kpi_glossary.json")
SHEET_NAME = "00_50CHITIEU"

# Marker để đánh dấu chỉ tiêu CHƯA chốt được nguồn (analyst/QA cần thận trọng).
FOLLOWUP_MARKERS = [
    "cần chốt", "chưa cập nhật", "chưa có", "check lại", "hỏi lại",
    "cần trao đổi", "note:", "cần hỏi", "trao đổi", "hỏi lại", "chốt lại",
]


def _norm(s) -> str:
    s = str(s or "").strip().lower().replace("đ", "d")  # NFKD KHÔNG tách đ(U+0111)
    s = unicodedata.normalize("NFKD", s)
    return " ".join("".join(c for c in s if not unicodedata.combining(c)).split())


# Nhãn header -> khoá chuẩn (khớp theo _norm, bền với đổi thứ tự/typo cột).
_COL_ALIASES = {
    "stt": "stt",
    "nhom chi tieu": "nhom_con", "nhom chi tieu (theo html)": "nhom_con",
    "chi tieu": "chi_tieu", "chi tieu (ten dung tren html)": "chi_tieu",
    "gia tri mau": "gia_tri_mau", "gia tri mau tren html": "gia_tri_mau",
    "trang thai den": "trang_thai_den", "trang thai den mau": "trang_thai_den",
    "cong thuc": "cong_thuc", "cong thuc / logic": "cong_thuc",
    "dvt": "don_vi",
    "nguon du lieu": "nguon_du_lieu",
    "sheet nhap lieu": "sheet_nhap",
    "cot nguon": "cot_nguon", "cot nguon trong sheet": "cot_nguon",
    "chieu phan tich": "chieu_phan_tich",
    "nguong vang": "nguong_vang",
    "nguong do": "nguong_do",
    "chieu danh gia": "chieu_danh_gia",
}


def _find_header(ws):
    """Tìm dòng header (dòng có ô 'Chỉ tiêu' + 'Công thức'); trả (row_idx, {col_idx: key})."""
    for r in range(1, min(ws.max_row, 6) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        keys = {}
        for ci, v in enumerate(vals, start=1):
            k = _COL_ALIASES.get(_norm(v))
            if k:
                keys[ci] = k
        if "chi_tieu" in keys.values() and "cong_thuc" in keys.values():
            return r, keys
    raise RuntimeError("Không tìm thấy dòng header trong 00_50CHITIEU")


def build_glossary(path=None, sheet=SHEET_NAME):
    path = path or TEMPLATE_XLSX
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    hdr_row, colmap = _find_header(ws)

    records = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        rec = {"row": r}
        for ci, key in colmap.items():
            v = ws.cell(r, ci).value
            rec[key] = "" if v is None else str(v).strip()
        chi_tieu = rec.get("chi_tieu", "")
        # Bỏ dòng tiêu đề nhóm ('▶ ...') / dòng rỗng (không có tên chỉ tiêu thật).
        if not chi_tieu or chi_tieu.startswith("▶"):
            continue
        # Trường tương thích với glossary_lookup (qa_server đọc các khoá này).
        rec.setdefault("nhom_bao_cao", "")
        rec.setdefault("nhom_con", rec.get("nhom_con", ""))
        rec.setdefault("chieu_phan_tich", rec.get("chieu_phan_tich", ""))
        rec.setdefault("cong_thuc", rec.get("cong_thuc", ""))
        rec.setdefault("nguon_du_lieu", rec.get("nguon_du_lieu", ""))
        # 'canh_bao_do' cũ ~ ngưỡng đỏ (để glossary_lookup vẫn khớp từ khoá cảnh báo).
        rec["canh_bao_do"] = rec.get("nguong_do", "")
        hay = " ".join([rec.get("chi_tieu", ""), rec.get("cong_thuc", ""),
                        rec.get("nguon_du_lieu", ""), rec.get("chieu_phan_tich", "")]).lower()
        rec["needs_followup"] = any(m in hay for m in FOLLOWUP_MARKERS)
        rec["wired"] = bool(rec.get("sheet_nhap")) and not rec["needs_followup"]
        records.append(rec)

    for idx, rec in enumerate(records, start=1):
        rec["id"] = idx
    return records


def main():
    if not os.path.exists(TEMPLATE_XLSX):
        print(f"Không tìm thấy template vàng: {TEMPLATE_XLSX}", file=sys.stderr)
        sys.exit(1)
    records = build_glossary()
    with open(OUT_PATH, "w", encoding="utf-8") as fh:
        json.dump(records, fh, ensure_ascii=False, indent=2)
    n_follow = sum(1 for r in records if r["needs_followup"])
    n_wired = sum(1 for r in records if r["wired"])
    print(f"Đã ghi {OUT_PATH}: {len(records)} chỉ tiêu "
          f"({n_wired} đã nối sheet, {n_follow} cần chốt nguồn). Nguồn: {SHEET_NAME} trong Template_chuan.")


if __name__ == "__main__":
    main()

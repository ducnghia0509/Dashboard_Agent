# -*- coding: utf-8 -*-
"""Đối chiếu SỐ agent trả về với SỐ trong brief.

Khác hẳn `run_chutich_eval.py`: bộ đó đo ĐỊNH DẠNG (có số không, có dòng cập nhật không, có từ
chối đúng không) — một câu trả lời bịa số vẫn qua được nếu trình bày đẹp. Bộ này đo ĐÚNG/SAI của
chính con số, là thứ Chủ tịch dùng để ra quyết định.

Giá trị mong đợi đọc từ brief LÚC CHẠY, không hardcode: daemon dựng lại brief mỗi giờ nên số
thay đổi liên tục; hardcode là tự tạo ra một nguồn sự thật thứ hai rồi lại lệch.

Chạy:
    .venv/bin/python scripts/check_chutich_accuracy.py
    .venv/bin/python scripts/check_chutich_accuracy.py --ids TIEN_TONG VAY_TONG
"""
import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta, timezone

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
for _p in (_ROOT, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from scripts.run_chutich_eval import ask  # noqa: E402  — dùng lại đúng đường gọi agent

BRIEF = os.path.normpath(os.path.join(_ROOT, "..", "Tài liệu", "BRIEF_CHUTICH.xlsx"))
OUT_DIR = os.path.join(_ROOT, "eval", "chutich")
TZ = timezone(timedelta(hours=7))

# Sai số cho phép. Agent được phép làm tròn (384.849 -> 384,8 hoặc 385) nên đối chiếu theo tỷ lệ
# chứ không tuyệt đối. 2% đủ rộng cho làm tròn 1 chữ số thập phân, vẫn đủ chặt để bắt lệch thật
# (vụ tiền T08: 1,157 so với 267,681 là lệch 99,6%).
TOL = 0.02


def _norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _rows(sheet):
    wb = openpyxl.load_workbook(BRIEF, data_only=True)
    try:
        return [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    finally:
        wb.close()


def _cell_by_key(sheet, key, col):
    """Lấy ô ở cột `col` của dòng có cột A khớp `key`."""
    for r in _rows(sheet):
        if str(r[0] or "").strip() == key:
            return r[col]
    return None


def _first_data_row(sheet, skip_header=1):
    rows = _rows(sheet)
    return rows[skip_header] if len(rows) > skip_header else []


# Mỗi phép kiểm: id, câu hỏi (đúng nguyên văn như trong questions.json), hàm lấy số mong đợi,
# và (tuỳ chọn) danh sách chuỗi BẮT BUỘC xuất hiện — vì nhiều câu sai ở TÊN chứ không ở số.
CHECKS = [
    {"id": "NO_QUA_HAN", "q": "Nợ quá hạn bao nhiêu?",
     "so": lambda: _cell_by_key("L5_CONGNO", "QUÁ HẠN", 4)},
    {"id": "NO_180", "q": "Có khoản nợ nào trên 180 ngày?",
     "so": lambda: _cell_by_key("L5_CONGNO", "NGUY CƠ MẤT VỐN", 4)},
    # Câu này hỏi TỔNG phải thu, nên đối chiếu được với dòng TỔNG. Bản đầu tôi gắn dòng TỔNG vào
    # câu "Công nợ gồm những khách hàng nào?" — câu đó hỏi DANH SÁCH, agent liệt kê khách là
    # đúng, chấm trượt là oan.
    {"id": "CONGNO_TONG", "q": "Có doanh thu nào chưa thu tiền?",
     "so": lambda: _cell_by_key("L5_CONGNO", "TỔNG", 4)},
    # Khớp theo đoạn ĐẶC TRƯNG, bỏ dấu: agent viết "Công ty CP VinFast Việt Nam" trong khi brief
    # ghi "Công ty Cổ phần VinFast Việt Nam". Viết tắt "Cổ phần" -> "CP" là bình thường và với
    # Chủ tịch còn dễ đọc hơn; đòi khớp nguyên văn là đo cách gõ chữ, không đo tính đúng.
    {"id": "TOP_KHACH", "q": "Top 20 khách nợ nhiều nhất.",
     "so": lambda: _cell_by_key("L5_CONGNO", "Top khách nợ #1", 4),
     "text": lambda: ["vinfast viet nam"]},
    {"id": "TIEN_TONG", "q": "Tiền đang nằm ở đâu?",
     "so": lambda: _cell_by_key("L4_DONGTIEN", "TIỀN — TỔNG", 2)},
    # PHÉP KIỂM NGƯỢC. Brief chỉ có TỔNG dư nợ vay, KHÔNG có lịch đáo hạn từng khoản. Nên câu
    # "khoản vay nào sắp đến hạn" phải được TỪ CHỐI, và tuyệt đối không được đem tổng dư nợ ra
    # trả lời — số đúng của một câu hỏi khác vẫn là số sai. Agent đang làm đúng (11/08); test này
    # để nếu sau đó nó bắt đầu bịa thì bắt được ngay.
    {"id": "VAY_KHONG_LAY_TONG", "q": "Khoản vay nào sắp đến hạn?",
     "so": None, "khong_duoc_co": lambda: _cell_by_key("L4_DONGTIEN", "VAY — TỔNG", 2)},
    {"id": "THU_HET", "q": "Nếu thu hết công nợ thì dòng tiền tăng bao nhiêu?",
     "so": lambda: _cell_by_key("L4_DONGTIEN", "NẾU THU HẾT CÔNG NỢ", 2)},
    {"id": "THU_30", "q": "Nếu thu được 30% công nợ thì dòng tiền cải thiện bao nhiêu?",
     "so": lambda: _cell_by_key("L4_DONGTIEN", "NẾU THU 30% CÔNG NỢ", 2)},
    {"id": "DT_CAO_NHAT", "q": "Đơn vị nào đóng góp doanh thu nhiều nhất?",
     "so": lambda: _first_data_row("L2_DOANHTHU")[4],
     "text": lambda: [str(_first_data_row("L2_DOANHTHU")[0] or "")]},
    {"id": "XEP_HANG_1", "q": "Xếp hạng toàn bộ công ty từ tốt đến xấu.",
     "so": lambda: _first_data_row("L8_XEPHANG")[5],
     "text": lambda: [str(_first_data_row("L8_XEPHANG")[1] or "")]},
]

# Token số trong câu trả lời: '384.849', '384,849', '1.499,549', '23,67'
_RE_NUM = re.compile(r"\d[\d.,]*\d|\d")


def _ung_vien(token: str):
    """Một token -> các cách đọc có thể. Tiếng Việt dùng ',' làm thập phân, tài liệu kỹ thuật
    lại dùng '.', và agent trộn cả hai. Sinh mọi cách đọc rồi để phép so tự chọn."""
    out = set()
    t = token.strip().rstrip(".,")
    if not t:
        return out
    for bien in (t, t.replace(".", ""), t.replace(",", ""),
                 t.replace(".", "@").replace(",", ".").replace("@", ""),
                 t.replace(",", "")):
        try:
            out.add(float(bien))
        except ValueError:
            pass
    # '1.499,549' -> bỏ '.' nghìn, ',' thành thập phân
    if "." in t and "," in t:
        try:
            out.add(float(t.replace(".", "").replace(",", ".")))
        except ValueError:
            pass
    return out


def _co_so(answer: str, mong_doi: float, tol=TOL):
    """Trong câu trả lời có con số nào khớp `mong_doi` (theo tỷ lệ tol) không."""
    if mong_doi is None:
        return None, []
    muc = abs(mong_doi) * tol
    gan = []
    for tok in _RE_NUM.findall(answer):
        for v in _ung_vien(tok):
            if abs(v - mong_doi) <= max(muc, 0.001):
                return True, []
            if mong_doi and 0.5 <= abs(v / mong_doi) <= 2 and v != 0:
                gan.append(v)
    return False, sorted(set(gan), key=lambda x: abs(x - mong_doi))[:4]


def main():
    ap = argparse.ArgumentParser(description="Đối chiếu số agent trả về với số trong brief")
    ap.add_argument("--ids", nargs="+", help="chỉ chạy các phép kiểm chỉ định")
    ap.add_argument("--model", help="ép model cho lượt chạy")
    a = ap.parse_args()

    checks = [c for c in CHECKS if not a.ids or c["id"] in set(a.ids)]
    stamp = datetime.now(TZ).strftime("%Y%m%d-%H%M%S")
    meta_as_of = _cell_by_key("_META", "cap_nhat_hien_thi", 1)
    print(f"Đối chiếu {len(checks)} phép kiểm · brief {meta_as_of} · sai số {TOL * 100:.0f}%\n")

    results, dat = [], 0
    for i, c in enumerate(checks, 1):
        mong_doi = c["so"]() if c.get("so") else None
        cam = c["khong_duoc_co"]() if c.get("khong_duoc_co") else None
        ans, secs, err = ask(c["q"], f"acc-{stamp}-{c['id']}", model=a.model)
        loi = []
        if err:
            loi.append(f"lỗi gọi agent: {err}")
        else:
            if c.get("so"):
                khop, gan = _co_so(ans, mong_doi)
                if khop is None:
                    loi.append("brief không có giá trị mong đợi (mốc đổi tên?)")
                elif not khop:
                    loi.append(f"không thấy số {mong_doi}"
                               + (f" · số gần nhất agent nói: {gan}" if gan
                                  else " · không có số nào gần"))
            if cam is not None:
                co_cam, _ = _co_so(ans, cam)
                if co_cam:
                    loi.append(f"LỖI NẶNG: đem số {cam} (của câu hỏi khác) ra trả lời")
            for t in (c.get("text", lambda: [])() or []):
                if t and _norm(t) not in _norm(ans):
                    loi.append(f"thiếu chuỗi bắt buộc: '{t}'")
        ok = not loi
        dat += ok
        results.append({**{k: c[k] for k in ("id", "q")}, "mong_doi": mong_doi,
                        "seconds": round(secs, 2), "pass": ok, "fails": loi, "answer": ans})
        print(f"  [{i:2d}/{len(checks)}] {c['id']:12s} {'ĐÚNG' if ok else 'SAI '} "
              f"{secs:6.1f}s  mong đợi {mong_doi}")
        for l in loi:
            print(f"                 - {l}")

    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"doichieu-{stamp}.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump({"at": stamp, "brief_as_of": meta_as_of, "tol": TOL, "results": results},
                  fh, ensure_ascii=False, indent=2)
    print(f"\nĐã lưu {out}")
    print("=" * 70)
    print(f"KHỚP SỐ: {dat}/{len(checks)}")
    print("=" * 70)
    return 0 if dat == len(checks) else 1


if __name__ == "__main__":
    sys.exit(main())

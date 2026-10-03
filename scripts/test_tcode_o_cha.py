# -*- coding: utf-8 -*-
"""`derive_hqkd_ngay._tcode_cong_tu_con` — khi nào tin ô cha, khi nào cộng từ dòng con (Trạm sạc).

Chạy: python scripts/test_tcode_o_cha.py

Hai ca thật:
  · 21/09/2026: ô cha T203/T200 = 0 nhưng dòng con có 5.278.400 -> phải CỘNG TỪ CON.
  · 30/09/2026 bản sửa 16:48: ô T201 ghi CỨNG 400.016.106, bốn dòng con trống/0 -> phải GIỮ Ô CHA
    (trước 03/10/2026 bị thay bằng Σ con = 0, mất 400 tr giá vốn, T200 còn 1,956 thay vì 2,356 tỷ).
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
import derive_hqkd_ngay as d  # noqa: E402

HDR = [None, "Mã số", "Chỉ tiêu", "D30"]


def _rows(t201, con201, t203, con203):
    return [
        HDR,
        [None, "T100", "DOANH THU TRẠM SẠC", 0],
        [None, "T101", "Doanh thu bán hàng", 0],
        [None, "T200", "CHI PHÍ TRẠM SẠC", 0],
        [None, "T201", "Gía vốn hàng bán", t201],
        ["- Giá vốn Trạm sạc", None, "Chi phí giá vốn hoa hồng", con201[0]],
        ["- Giá vốn Trạm sạc", None, "Chi phí giá vốn bán trụ", con201[1]],
        ["Chi phí tài chính", "T202", "Chi phí tài chính", 113021],
        [None, "T203", "Chi phí bán hàng", t203],
        ["- CP bán hàng", None, "Chi phí nhân viên", con203],
        [None, "T204", "Chi phí quản lý doanh nghiệp", 264566198],
        [None, "T300", "LỢI NHUẬN TRẠM SẠC", 0],
    ]


loi = []


def kiem(ten, dung):
    print(("  ĐẠT  " if dung else "  SAI  ") + ten)
    if not dung:
        loi.append(ten)


moi, _ = d._tcode_cong_tu_con(_rows(400016106, (None, 0), 1691769900, 1691769900))
print("1. Ô cha ghi cứng, dòng con trống/0 (30/09 bản 16:48)")
kiem("T201 giữ 400.016.106", moi["T201"][1] == 400016106)
kiem("T200 = 2.356.465.225", moi["T200"][1] == 400016106 + 113021 + 1691769900 + 264566198)
kiem("T300 = -T200 (doanh thu 0)", moi["T300"][1] == -moi["T200"][1])

moi, lech = d._tcode_cong_tu_con(_rows(0, (None, 0), 0, 5278400))
print("2. Ô cha 0, dòng con có số (21/09)")
kiem("T203 cộng từ con = 5.278.400", moi["T203"][1] == 5278400)
kiem("T203 báo lệch cha/con", "T203" in lech)

print()
if loi:
    print(f"THẤT BẠI {len(loi)} ca:")
    for x in loi:
        print("   -", x)
    sys.exit(1)
print("TẤT CẢ ĐẠT")

# -*- coding: utf-8 -*-
"""Hook map tên đơn vị -> cost_center/công ty: BIẾN THỂ "<đơn vị>_ HL" và cái bẫy bản "(61)".

Chạy: python scripts/test_cc_bien_the_ten.py

Vì sao có test này (24/08/2026): báo cáo công nợ SRVF T5/T6 ghi tên đơn vị là "Showroom Uông Bí_
HL"/"_ CP" — nhóm bán hàng nội bộ của CÙNG showroom Uông Bí, nhưng không có trong danh mục cost
center nên hook trả `_khong_map`; spec khai `giu_khi_khong_map` nên 38 dòng dư nợ THẬT (14,10 tỷ)
vào DB với cost_center/cong_ty rỗng. Tổng khối không hụt, nên không ai thấy — cái vỡ là chúng
thành MỘT NHÓM SNAPSHOT RIÊNG (`_rows` chọn ngày chốt theo (cong_ty, khoi)): nhóm rỗng đứng mãi ở
30/06 nên màn Quản lý phải thu ở view nhiều kỳ cộng số dư T08 với phần sót T06.

Test khoá HAI chiều, vì cách vá dễ sai theo cả hai:
  · biến thể có `_` PHẢI quy về đơn vị gốc;
  · bản "(61)" (pháp nhân Xanh Vĩnh Phúc) PHẢI KHÔNG bị quy về bản gốc (Thịnh Cường). Nếu ai đó
    vá bằng cách khớp tiền tố SAU `_nd` thì `campha61` khớp `campha` và tiền của XVP chảy sang TC
    — không có test này thì lỗi đó im lặng hoàn toàn.
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..")))
sys.path.insert(0, _HERE)

import spec_extract as se  # noqa: E402

loi = []


def kiem(ten, cc, cty, ghi_chu=""):
    got = se._cc_showroom(ten)
    cho_doi = {"cost_center": cc, "cong_ty": cty} if cc else {"_khong_map": ten}
    thuc = ({k: got.get(k) for k in ("cost_center", "cong_ty")} if cc
            else {"_khong_map": got.get("_khong_map")})
    dau = "OK " if thuc == cho_doi else "SAI"
    if thuc != cho_doi:
        loi.append(f"{ten!r}: chờ {cho_doi}, nhận {thuc}")
    print(f"  [{dau}] {ten:32} -> {thuc}  {ghi_chu}")


print("1. Biến thể có dấu '_' quy về đơn vị gốc")
kiem("Showroom Uông Bí", "UB_SR", "VFQN", "bản gốc, phải giữ nguyên")
kiem("Showroom Uông Bí_ HL", "UB_SR", "VFQN", "38 dòng T5/T6")
kiem("Showroom Uông Bí_ CP", "UB_SR", "VFQN", "38 dòng T5/T6")
kiem("Showroom Uông Bí_ VP", "UB_SR", "VFQN", "biến thể CHƯA từng gặp")

print("\n2. Bản '(61)' KHÔNG được quy về bản gốc — khác PHÁP NHÂN")
kiem("Vinfast Cẩm Phả (61)", "CP_SR_61", "XVP", "Xanh Vĩnh Phúc, KHÔNG phải TC")
kiem("Showroom Cẩm Phả", "CP_SR", "TC", "Thịnh Cường")
kiem("Vinfast Hạ Long (61)", "HL_SR_61", "XVP")
kiem("Showroom Hạ Long", "HL_SR", "TC")

print("\n3. Tên KHÔNG có trong danh mục vẫn phải BÁO, không được im lặng nhận bừa")
kiem("Showroom Không Tồn Tại_ XX", None, None, "có '_' nhưng gốc cũng không có trong master")
kiem("CHI NHÁNH VINFAST HÀ NỘI", None, None, "không thuộc showroom nào")

print("\n4. Alias khai ở hook vẫn chạy")
kiem("Vinfast B2B", "B2B_SR", "TC", "đội bán B2B tập trung")

print()
if loi:
    print(f"THẤT BẠI {len(loi)} ca:")
    for x in loi:
        print("   -", x)
    sys.exit(1)
print("TẤT CẢ ĐẠT")

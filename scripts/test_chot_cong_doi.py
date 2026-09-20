# -*- coding: utf-8 -*-
"""Chốt CỘNG ĐÔI TRONG MỘT FILE (`_kiem_cong_doi`) — nổ khi dòng tiêu đề đơn vị lọt vào dữ liệu.

Chạy: python scripts/test_chot_cong_doi.py

Vì sao có test này (31/08/2026): kế toán đổi bố cục khối I của `Baocaodoanhthungay` — dòng tên
xưởng trước ở `cột A = "3S có đồng sơn" · cột B = "Ocean Park"` (KHÔNG mã), nay mang luôn mã
`B100`. Quy tắc `ngu_canh_dong` "dòng tên xưởng" đòi cột A không khớp mẫu mã nên không nổ nữa:
cả khối I mất cost_center VÀ chính dòng đó thành một dòng số nằm chung xô cấp khối với dòng tổng
của file -> doanh thu XDV T8 hiện 85,098 tỷ trong khi file ghi 42,549 tỷ, đúng gấp đôi.

Không một lớp chống trùng nào cũ bắt được (xem docstring `_kiem_cong_doi`), nên chốt này là lưới
duy nhất. Test khoá BA chiều, vì cách vá dễ sai theo cả ba:
  · trạng thái LÀNH của chính 3 spec phân cấp PHẢI đi qua (nhãn cấp khối là "TỔNG DOANH THU XDV",
    "Doanh thu công việc (XHĐ)", "Lệnh bảo hành (W)" — không cái nào là tên xưởng);
  · nhãn ĐƠN VỊ đứng ở cấp khối PHẢI bị chặn, dù nằm ở dim1, dim2 hay dim3;
  · spec PHẲNG (không khai `ngu_canh_dong` gán cost_center) PHẢI KHÔNG bị đụng tới — nhiều nguồn
    có tên đơn vị ở dim ở cấp khối một cách hợp lệ (HCNS_NS_PB để phòng/ban ở dim2, PTRA để tên
    đối tượng công nợ ở dim1). Siết cả chúng là chặn oan 92 spec còn lại.
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..")))
sys.path.insert(0, _HERE)

import spec_extract as se  # noqa: E402

SPEC_PHAN_CAP = {"ngu_canh_dong": [{"khi": [{"cot": "A", "dieu_kien": "khac_rong"}],
                                    "gan": {"cost_center": {"cot": "B", "chuan_hoa": "xdv"}}}]}
SPEC_PHANG = {"cot": {"cost_center": {"header": "Tên DVCS"}}}

loi = []


def kiem(ten, spec, recs, cho_no, ghi_chu=""):
    got = se._kiem_cong_doi(spec, recs)
    no = got is not None
    if no != cho_no:
        loi.append(f"{ten}: chờ {'NỔ' if cho_no else 'im'}, nhận {'NỔ' if no else 'im'} ({got})")
    print(f"  [{'OK ' if no == cho_no else 'SAI'}] {ten:52} -> {'NỔ' if no else 'im'}  {ghi_chu}")


print("0. master_data đọc được (nếu không, chốt tự tắt và test 2 vô nghĩa)")
ten_dv = se._moi_ten_don_vi()
print(f"  [{'OK ' if len(ten_dv) > 1 else 'SAI'}] {len(ten_dv) - 1} khoá tên đơn vị")
if len(ten_dv) <= 1:
    loi.append("không nạp được master_data -> không kiểm được chốt")

print("\n1. Trạng thái LÀNH của file doanh thu XDV — phải đi qua")
lanh = [{"dim1": "B100", "dim3": "TỔNG DOANH THU XDV", "cost_center": None},
        {"dim1": "B110", "dim3": "Doanh thu công việc (XHĐ)", "cost_center": None},
        {"dim1": "B120", "dim3": "Doanh thu phụ tùng (XHĐ)", "cost_center": None},
        {"dim1": "B110", "dim3": "Doanh thu công việc (XHĐ)", "cost_center": "OCP_XDV"},
        {"dim1": "B120", "dim3": "Doanh thu phụ tùng (XHĐ)", "cost_center": "LB_XDV"}]
kiem("khối + xưởng, nhãn khối không phải tên xưởng", SPEC_PHAN_CAP, lanh, False, "bản đã sửa")
kiem("nhãn trạng thái RO ở cấp khối", SPEC_PHAN_CAP,
     [{"dim1": "Lệnh bảo hành (W)", "dim2": "Nghiệm thu đã XHĐ", "cost_center": None},
      {"dim1": "Lệnh bảo hành (W)", "dim2": "Nghiệm thu đã XHĐ", "cost_center": "HL_XDV"}],
     False, "xdv_doanhthu_ro / xdv_soluong_ro")

print("\n2. Nhãn ĐƠN VỊ đứng ở cấp khối — phải chặn, ở dim nào cũng vậy")
for cot in ("dim1", "dim2", "dim3"):
    kiem(f"tên xưởng ở {cot}, cost_center rỗng", SPEC_PHAN_CAP,
         [{"dim1": "B100", "dim3": "TỔNG DOANH THU XDV", "cost_center": None},
          {"dim1": "B110", "dim3": "x", "cost_center": "OCP_XDV"},
          {cot: "Ocean Park", "cost_center": None}],
         True, "đúng lỗi 31/08/2026")
kiem("KHÔNG dòng nào nhận được cost_center", SPEC_PHAN_CAP,
     [{"dim1": "B110", "dim3": "Doanh thu công việc (XHĐ)", "cost_center": None}] * 3,
     True, "quy tắc tiêu đề chết hẳn")

print("\n3. Spec PHẲNG không bị đụng tới")
kiem("tên đơn vị ở dim2, spec không khai ngu_canh_dong", SPEC_PHANG,
     [{"dim2": "Depot Phú Thọ", "cost_center": None}], False, "HCNS_NS_PB")
kiem("tên đơn vị ở dim1, spec không khai ngu_canh_dong", SPEC_PHANG,
     [{"dim1": "DEPOT TUYÊN QUANG", "cost_center": None}], False, "PTRA")
kiem("recs rỗng", SPEC_PHAN_CAP, [], False, "file không đọc được -> để lớp khác xử")

print()
if loi:
    print(f"THẤT BẠI {len(loi)} ca:")
    for x in loi:
        print("   -", x)
    sys.exit(1)
print("TẤT CẢ ĐẠT")

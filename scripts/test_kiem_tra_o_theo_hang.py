# -*- coding: utf-8 -*-
"""Chốt LAYOUT kiểm theo HÀNG (`kiem_tra_o` với khoá `hang`) — mở khoá ca nguồn chèn thêm cột.

Chạy: python scripts/test_kiem_tra_o_theo_hang.py

## Vì sao có test này (sự cố thật 14-15/09/2026)

Cyber chèn thêm cột "Thanh toán VinPoint" vào giữa bảng kê hoá đơn bán xe: file từ 56 lên 57
cột, mọi cột từ P trở đi dịch phải một ô. `vhkd_kqkd_auto` địa chỉ cột theo TÊN HEADER nên việc
đọc số không hề gì — nhưng cái neo `{"o": "R8", "bang": "Ngày hóa đơn"}` nay trỏ vào "Tư vấn bán
hàng", và chốt từ chối CẢ FILE.

Hậu quả: 61 hoá đơn ngày 14/09 không vào DB. Bản tin 07:30 sáng hôm sau báo XHĐ = 0 cho cả 9
showroom, đỏ toàn bảng — một báo động giả gửi thẳng tới giám đốc khối, và người phụ trách phải
tự chụp số thật gửi lại để đính chính.

Chốt đã chặt hơn thứ nó cần chặn. Việc nó PHẢI phát hiện là "file này khác LOẠI báo cáo"; đòi
nhãn nằm đúng ô thì mỗi lần nguồn chèn một cột vô hại là mất trắng một ngày dữ liệu.

## Ba chiều test khoá, vì cách vá dễ sai theo cả ba

1. Chèn cột vô hại -> PHẢI đi qua (đây là ca đã hỏng).
2. File KHÁC LOẠI (thiếu nhãn mốc) -> PHẢI vẫn bị chặn. Nới quá tay là mất luôn lưới an toàn
   vốn đã cứu vụ công nợ T1 (header dòng 5, nạp bừa đẻ ra 9,86 tỷ vô nghĩa).
3. Khai kiểu cũ `{"o": ...}` -> PHẢI giữ nguyên hành vi. 90+ spec đang dùng nó.
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..")))
sys.path.insert(0, _HERE)

import openpyxl  # noqa: E402

import spec_extract as se  # noqa: E402

_loi = []


def _kiem(ten, thuc, mong):
    dat = thuc == mong
    print(f"  [{'OK ' if dat else 'SAI'}] {ten}: {thuc!r}")
    if not dat:
        _loi.append(ten)


def _wb(headers, tmp):
    """Dựng file có header ở DÒNG 8, dữ liệu từ dòng 9 — đúng khuôn nguồn Cyber."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for j, h in enumerate(headers, start=1):
        ws.cell(row=8, column=j, value=h)
    wb.save(tmp)
    return tmp


def _chay(spec, path):
    """-> (số dòng, cảnh báo đầu tiên)."""
    recs, warn = se._extract_vung(spec, path)
    return len(recs), (warn[0] if warn else "")


BINH_THUONG = ["Tên DVCS", "Mã Hợp đồng", "Ngày hợp đồng", "Điểm tích lũy",
               "Tư vấn bán hàng", "Ngày hóa đơn", "Số lượng"]
#: Đúng cách Cyber đã đổi: chèn MỘT cột vào giữa, đẩy mọi cột sau sang phải.
CHEN_COT = ["Tên DVCS", "Mã Hợp đồng", "Ngày hợp đồng", "Thanh toán VinPoint",
            "Điểm tích lũy", "Tư vấn bán hàng", "Ngày hóa đơn", "Số lượng"]
KHAC_LOAI = ["Mã số", "Lũy kế", "Kỳ này", "Showroom OceanPark"]


def _spec(kiem):
    return {"nguon": {"sheet": {"so": 0}}, "header": {"dong": 8}, "dong_bat_dau": 9,
            "kiem_tra_o": kiem,
            "cot": {"cost_center": {"header": "Tên DVCS", "bat_buoc": False}},
            "ban_ghi": "moi_dong"}


def main():
    import tempfile
    d = tempfile.mkdtemp()
    f_thuong = _wb(BINH_THUONG, os.path.join(d, "thuong.xlsx"))
    f_chen = _wb(CHEN_COT, os.path.join(d, "chen.xlsx"))
    f_khac = _wb(KHAC_LOAI, os.path.join(d, "khac.xlsx"))

    theo_hang = _spec([{"hang": 8, "bang": "Tên DVCS"},
                       {"hang": 8, "bang": "Ngày hóa đơn"},
                       {"hang": 8, "bang": "Số lượng"}])
    theo_o = _spec([{"o": "F8", "bang": "Ngày hóa đơn"}])

    print("1. Chèn cột vô hại — ca đã hỏng ngày 14/09")
    _kiem("theo HÀNG: file bố cục thường đi qua", _chay(theo_hang, f_thuong)[1], "")
    _kiem("theo HÀNG: file CHÈN CỘT vẫn đi qua", _chay(theo_hang, f_chen)[1], "")
    n, w = _chay(theo_o, f_chen)
    _kiem("theo Ô: file chèn cột BỊ CHẶN (hành vi cũ, giữ nguyên)",
          w.startswith("BỎ QUA — layout khác spec: ô F8"), True)

    print("\n2. File khác loại vẫn phải bị chặn")
    n, w = _chay(theo_hang, f_khac)
    _kiem("theo HÀNG: file khác loại bị chặn", w.startswith("BỎ QUA — layout khác spec: hàng 8"), True)
    _kiem("   cảnh báo nêu đúng nhãn còn thiếu", "Tên DVCS" in w, True)

    print("\n3. Khai kiểu cũ không đổi hành vi")
    _kiem("theo Ô: file bố cục thường vẫn đi qua", _chay(theo_o, f_thuong)[1], "")

    print("\n4. Hai spec nguồn SR đã chuyển sang kiểu HÀNG")
    import json
    for ten in ("vhkd_kqkd_auto", "vhkd_hopdong"):
        sp = json.load(open(os.path.join(_HERE, "..", "extract_specs", f"{ten}.json"),
                            encoding="utf-8"))
        _kiem(f"{ten}: mọi mốc khai theo `hang`",
              all("hang" in k for k in sp["kiem_tra_o"]), True)

    print("\n" + ("== TẤT CẢ ĐẠT ==" if not _loi else f"== {len(_loi)} SAI: {_loi} =="))
    sys.exit(1 if _loi else 0)


if __name__ == "__main__":
    main()

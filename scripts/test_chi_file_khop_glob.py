# -*- coding: utf-8 -*-
"""Khoá `nguon.chi_file_khop_glob` của `spec_extract.specs_for_path`.

Chạy: python scripts/test_chi_file_khop_glob.py

Sự cố 06/10/2026: `xtai_kehoach_gt/_sl` dò header theo nhãn 'Khối' — nhãn có ở sheet KHDT của
mọi khối, nên autofill file 1.SR/6.XVP/7.AnTX cũng ghi vào XTAI_KH (T10 phồng 219,7 -> 1.862 tỷ).
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)

import spec_extract as se  # noqa: E402

KH = os.path.join(se.REPORTS_DIR, "KEHOACH", "baocaokehoachthang")
_loi = []


def _kiem(ten, dat):
    print(f"  [{'OK ' if dat else 'SAI'}] {ten}")
    if not dat:
        _loi.append(ten)


def _ids(ten_file):
    return {s["id"] for s in se.specs_for_path(os.path.join(KH, ten_file))}


def main():
    sr = _ids("1.SR.OO.M.202608.Kehoachthang.xlsx")
    ht = _ids("5.HT.OO.M.202608.Kehoachthang.xlsx")
    xtai = {"xtai_kehoach_gt", "xtai_kehoach_sl", "xtai_kehoach_ct_gt", "xtai_kehoach_ct_sl"}
    _kiem("spec Xe tải KHÔNG chạy trên file SR", not (sr & xtai))
    _kiem("spec Xe tải VẪN chạy trên file HT", xtai <= ht)
    _kiem("spec không khai cờ vẫn chọn theo thư mục (vhkd_kehoach_gt có ở file SR)",
          "vhkd_kehoach_gt" in sr)
    _kiem("spec không khai cờ vẫn chọn theo thư mục kể cả file lệch glob (vhkd_kehoach_gt ở file HT)",
          "vhkd_kehoach_gt" in ht)
    _kiem("tên viết hoa/thường lệch vẫn khớp", xtai <= _ids("5.HT.OO.M.202608.KEHOACHTHANG.XLSX"))
    print("ĐẠT" if not _loi else f"HỎNG {len(_loi)} ca")
    sys.exit(1 if _loi else 0)


if __name__ == "__main__":
    main()

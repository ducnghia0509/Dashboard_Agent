# -*- coding: utf-8 -*-
"""Dòng công nợ KHÔNG CÓ TÊN đối tượng phải vào tổng (14/10/2026). Chạy: .venv/bin/python scripts/test_sodu_dong_khong_ten.py
Trước đây `_facts_congno` bỏ dòng rỗng tên -> PTHU_D hụt so với tổng file (XDV 12/10: 15,4 triệu; 28/09: 1,11 tỷ)."""
import os
import sys

os.environ.setdefault("BACKEND_PATH", "/home/itadmin/AI_Dashboard_QT/DashBoard_AI/backend")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import derive_sodu_tudong_ngay as d  # noqa: E402


def dong(ten, cuoi, dau=0.0):
    return {"amount": cuoi, "pl": {"ten_doi_tuong": ten, "ma_doi_tuong": "", "no_dau_ky": dau,
                                    "co_dau_ky": 0.0, "ps_no": 0.0, "ps_co": 0.0}}


rows = [dong("Khách A", 10.0), dong("Khách A", 5.0), dong("Khách B", 3.0),
        dong(None, 0.015393), dong("", 0.5), dong("   ", 0.1)]
fs = d._facts_congno(rows, "PTHU_D", "no", "ten_doi_tuong", "ma_doi_tuong")
tong = sum(f[4] for f in fs)
ok = abs(tong - 18.615393) < 1e-9
ten = {f[1]: f[4] for f in fs}
ok &= abs(ten[d.KHONG_TEN_DOI_TUONG] - 0.615393) < 1e-9      # ba dòng thiếu tên gom về MỘT dòng riêng
ok &= abs(ten["Khách A"] - 15.0) < 1e-9 and len(fs) == 3
ok &= d._facts_congno([dong("Khách A", 1.0)], "PTHU_D", "no", "ten_doi_tuong", "ma_doi_tuong")[0][1] == "Khách A"
print("PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)

# -*- coding: utf-8 -*-
"""Khoá cách suy KỲ từ tên file — chạy: python scripts/test_ky_mo_ho.py

Chốt 2026-08-18: 3 file T10/T11/T12/2025 của SRVF ('B.1.TC.TCKT.M20250910/11/12') đều bị đọc
thành 2025-09 (regex ăn 6 số đầu, nuốt phần đuôi). Nạp vào là 3 dòng PTHU_TUOINO khác source_file
nằm chung một kỳ -> dashboard cộng 3 lần còn T10-T12 rỗng, im lặng. Nay trả None và bắt đổi tên.

Hai quy ước HỢP LỆ cũng có 8 chữ số liền nhau, KHÔNG được đụng: báo cáo NGÀY ('.D.20260801.' =
kỳ 2026-08, đuôi là ngày) và '.M.20260500.' của XDV (đuôi '00' không phải tháng).
"""
import os
import sys

_ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, _ROOT)
sys.path.insert(0, os.path.join(_ROOT, "scripts"))

from servers.template_filler import guess_period, ky_mo_ho  # noqa: E402
from sync_orchestrator import _guess_period, _ten_ky_mo_ho  # noqa: E402

CA = [
    ("B.1.TC.TCKT.M20250910.Baocaotuoino.Xlsx", None, True),
    ("B.1.TC.TCKT.M20250911.Baocaotuoino.Xlsx", None, True),
    ("B.1.TC.TCKT.M20250912.Baocaotuoino.Xlsx", None, True),
    ("B.1.TC.TCKT.M202509.Baocaotuoino.Xlsx", "2025-09", False),
    ("B.1.TC.TCKT.M202510.Baocaotuoino.Xlsx", "2025-10", False),   # tên ĐÚNG mong muốn
    ("B.1.TC.TCKT.M.202607.Baocaotuoino.Xlsx", "2026-07", False),
    ("B.9.TC.TCKT.D.20260801.Baocaotaichinhrieng.xlsx", "2026-08", False),   # báo cáo NGÀY
    ("B.2.TC.TCKT.M.20260500.BaocaocongnophaithuVF2026.xlsx", "2026-05", False),
    ("BC THÁNG 01 NĂM 2026.xlsx", "2026-01", False),
]

_fail = False
for ten, ky, mo_ho in CA:
    for nhan, f_ky, f_mo in (("template_filler", guess_period, ky_mo_ho),
                             ("sync_orchestrator", _guess_period, _ten_ky_mo_ho)):
        got_ky, got_mo = f_ky(ten), f_mo(ten)
        if got_ky != ky or got_mo != mo_ho:
            _fail = True
            print(f"[FAIL] {nhan}: {ten} -> kỳ={got_ky} (cần {ky}), mơ_hồ={got_mo} (cần {mo_ho})")
        else:
            print(f"[OK]   {nhan}: {ten} -> {got_ky or '—'}")

# Dự án: tháng KHÔNG zero-pad ('.20268.' = 2026-07/2026-08) — CHỈ sync_orchestrator có nhánh này
# (template_filler cố ý không, xem comment ở mỗi hàm). Khoá lại để không ai "đồng bộ" nhầm 2 bên.
for ten, ky in (("B.4.TC.TCKT.M.20267.Baocaotuoinophaithu.xlsx", "2026-07"),
                ("B.4.TC.TCKT.D.20268.Baocaohqkdngay.xlsx", "2026-08")):
    if _guess_period(ten) != ky or guess_period(ten) is not None:
        _fail = True
        print(f"[FAIL] '{ten}': sync={_guess_period(ten)} (cần {ky}), "
              f"filler={guess_period(ten)} (cần None)")
    else:
        print(f"[OK]   tháng 1 chữ số: {ten} -> sync {ky} · filler —")

print("\n== TẤT CẢ TEST PASS ==" if not _fail else "\n== CÓ TEST FAIL ==")
sys.exit(1 if _fail else 0)

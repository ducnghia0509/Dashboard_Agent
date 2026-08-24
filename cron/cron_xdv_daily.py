#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QTVH XƯỞNG DỊCH VỤ VINFAST (XDV) — chạy 19:00 giờ VN mỗi ngày.

Nuôi 5 màn `xdv0..xdv4` (Quản trị vận hành xưởng dịch vụ). Khung chung + 4 cái bẫy: xem
`cron_qtvh_core.py`. File này CHỈ khai báo nguồn.

MỐC 19:00 VN — MUỘN HƠN JOB SRVF (11:00) LÀ CỐ Ý, theo nhịp nộp thật (đo trên
`available_metadata.json` 24/08/2026): 5 nguồn KSCL tuần nộp 18:02–18:31 giờ VN, doanh thu ngày nộp
16:04. Chạy 11:00 như bên SRVF thì mỗi lượt đều trượt bản của chính ngày hôm đó.
MỐC CRONTAB THEO UTC (máy TZ=Etc/UTC, cron Ubuntu bỏ qua CRON_TZ khi tính lịch): 19:00 VN = 12:00 UTC.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_xdv_daily.py [--dry-run] [--env test|prod]

5 NGUỒN KSCL LÀ `anh_chup_ky`, KHÔNG PHẢI `luy_ke` — ĐỪNG ĐỔI. Mỗi file là dư nợ/tồn tại THỜI ĐIỂM
chốt tuần và từ 12/08/2026 mỗi tuần vào DB một ngày chốt riêng; cả 5 report_type đều nằm trong
`_SNAP_RT` nên màn hình tự lấy tuần mới nhất, còn bảng biến động cần cả chuỗi. Xoá bản cũ ở đây là
mất dữ liệu thật, không phải dọn trùng.

VÌ SAO KHÔNG CÓ `baocaohqkdngay` / `baocaotaichinhrieng`: cùng lý do như bên SRVF — xem docstring
`cron_srvf_daily.py`.
"""
import sys

import cron_qtvh_core as core

JOB = "xdv_daily"
NHAN = "Nguồn QTVH Xưởng dịch vụ Vinfast (XDV)"
SCHEDULE_VN = "19:00"

NGUON = [
    # ── ảnh chụp tuần (KSCL): nạp MỌI bản chưa có, KHÔNG xoá bản nào ────────────────────────
    {"company": "XDV", "rt": "baocaocongnoxhd", "che_do": core.ANH_CHUP_KY,
     "ten": "Công nợ đã xuất hoá đơn (tuần)"},
    {"company": "XDV", "rt": "baocaolsccyber", "che_do": core.ANH_CHUP_KY,
     "ten": "Lệnh sửa chữa Cyber (tuần)"},
    {"company": "XDV", "rt": "baocaolscdms", "che_do": core.ANH_CHUP_KY,
     "ten": "Lệnh sửa chữa tồn DMS (tuần)"},
    {"company": "XDV", "rt": "baocaolscchuaxhd", "che_do": core.ANH_CHUP_KY,
     "ten": "Lệnh sửa chữa chưa xuất hoá đơn (tuần)"},
    {"company": "XDV", "rt": "baocaolsctreodxbh", "che_do": core.ANH_CHUP_KY,
     "ten": "Treo đề xuất bảo hành (tuần)"},
    # ── tháng: 1 file/kỳ, luỹ kế sẵn theo ngày bên trong -> kéo lại đè chính nó ─────────────
    {"company": "XDV", "rt": "baocaodoanhthungay", "che_do": core.THANG,
     "ten": "Doanh thu + RO theo ngày"},
    {"company": "XDV", "rt": "baocaodoanhthukehoachngay", "che_do": core.THANG,
     "ten": "Kế hoạch doanh thu theo ngày"},
    {
        "company": "KEHOACH", "rt": "baocaokehoachthang", "che_do": core.THANG,
        "ten": "Kế hoạch tháng Xưởng dịch vụ",
        # `2.XDV.` = phần xưởng dịch vụ trong thư mục kế hoạch dùng chung (khớp `file_glob` của 6
        # spec xdv_kehoach_*). Không lọc là kéo trùng với job SRVF và cả 4 khối khác.
        "chi_lay": r"^2\.XDV\.",
    },
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

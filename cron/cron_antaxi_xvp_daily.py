#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QTVH AN TAXI + XANH VĨNH PHÚC — 3 lượt/ngày, xem khối MỐC GIỜ bên dưới.

Nuôi `at1..at5` (Quản trị vận hành An Taxi) và `xt0..xt5` (Xanh Vĩnh Phúc). Khung chung + các bẫy
của khung: xem `cron_qtvh_core.py`. File này CHỈ khai báo nguồn.

MỘT JOB CHO CẢ HAI KHỐI, KHÔNG TÁCH ĐÔI: mỗi khối chỉ có 2 thư mục nguồn, cùng một nhịp nộp, và
`cron_status` vốn ghi trạng thái THEO TỪNG NGUỒN (`expected` = danh sách company/report_type) nên
gộp job KHÔNG làm mất khả năng chỉ đúng khối nào nộp trễ. Tách đôi chỉ để thêm 6 dòng crontab nữa
cùng bắn một lúc.

BA LƯỢT/NGÀY: 05:00 · 16:45 · 17:15 giờ VN — cùng mốc với 3 job kia (thêm 29/08/2026).
Lượt PROD chạy sau TEST 5 phút — KHÔNG được bỏ: cùng phút thì hai lượt tranh khoá per-file
(servers/common/filelock.py), lượt sau bị 'skipped_lock' và MẤT HẲN một lượt nạp.
MỐC CRONTAB VIẾT THEO UTC (máy TZ=Etc/UTC, cron Ubuntu bỏ qua CRON_TZ khi TÍNH LỊCH — man 5
crontab, LIMITATIONS): 09:45 · 10:15 · 22:00 UTC. Lượt 05:00 VN nằm ở 22:00 UTC HÔM TRƯỚC.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_antaxi_xvp_daily.py [--dry-run] [--env test|prod]

TRƯỚC KHI CÓ JOB NÀY (đo 29/08/2026) 4 nguồn dưới đây kéo TAY hoàn toàn, và đã trôi:
  · `ANTAXI/baocaoqtvhngay` + `baocaoqtvhthang`: bản trên đĩa từ 28/08 16:34 VN, DB dừng ở 27/08.
  · `XANHVINHPHUC/baocaoqtvhthang`: bản trên đĩa từ **18/08** — trễ 11 ngày.
Đó là lý do job này tồn tại, không phải để chạy cho đủ bộ.

VÌ SAO KHÔNG CÓ `baocaohqkdngay` CỦA CẢ HAI ĐƠN VỊ: `cron_hqkdngay_daily.py` đã kéo (12 đơn vị
dùng CHUNG report_type đó) và đường `autofill` của nó gọi `spec_extract.run_for_path`, nên spec
`xvp_doanhthu_ngay` (XVP_DT_NGUON_D) cũng chạy theo — đã kiểm 29/08: file 29/08 trích ra 560 dòng,
DB có đúng 560. Khai lại ở đây là kéo trùng một file hai lần mỗi lượt.

VÌ SAO KHÔNG CÓ `baocaotaichinhrieng` / `baocaotuoino` / `baocaotaisancodinhcongcudungcu`: ba thư
mục đó là báo cáo TÀI CHÍNH THÁNG dùng chung cơ chế với MỌI đơn vị khác (đường derive_* cũ), không
riêng gì hai khối này. Tự động hoá chúng là một việc riêng, phải làm cho cả 12 đơn vị một lượt —
làm lẻ 2 đơn vị ở đây thì 10 đơn vị còn lại vẫn tay, mà lại tưởng là đã xong.
"""
import sys

import cron_qtvh_core as core

JOB = "antaxi_xvp_daily"
NHAN = "Nguồn QTVH An Taxi + Xanh Vĩnh Phúc"
SCHEDULE_VN = "05:00 · 16:45 · 17:15"   # 3 lượt/ngày. Ghi vào artifact cho agent giám sát khỏi
# hard-code mốc giờ — KHÔNG ai parse chuỗi này, chỉ hiển thị.

NGUON = [
    # ── An Taxi ────────────────────────────────────────────────────────────────────────────
    # 1 file/tháng, tên cố định theo tháng ('...D.202608.Baocaotonghop'), luỹ kế sẵn từng ngày
    # bên trong -> kéo lại là đè chính nó (idempotent theo source_file). Nuôi 6 spec ATX_*_D.
    {"company": "ANTAXI", "rt": "baocaoqtvhngay", "che_do": core.THANG,
     "ten": "QTVH An Taxi theo ngày"},
    {
        "company": "ANTAXI", "rt": "baocaoqtvhthang", "che_do": core.THANG,
        "ten": "QTVH An Taxi theo tháng",
        # `ca_nam` BẮT BUỘC: tên file là 'B.7.AAG.PKDVH.M.2026.BAOCAOTONGHOP.xlsx' — CÓ NĂM,
        # KHÔNG CÓ THÁNG. Bên quét trả `month=null` và `thang_tu_ten_file` cũng không đọc ra, nên
        # phép lọc `thang != month` (None != 8) sẽ loại file khỏi MỌI lượt kéo, IM LẶNG — đúng
        # lớp bug đã làm Dự án đứng số suốt một tháng (xem `cron_status.ky_cua_entry`). Một file
        # mang cả năm nên cứ kéo lại ở kỳ chính là đủ; 6 spec ATX_* tự bóc tháng từ trong sheet.
        "ca_nam": True,
    },
    {
        "company": "KEHOACH", "rt": "baocaokehoachthang", "che_do": core.THANG,
        "ten": "Kế hoạch tháng An Taxi",
        # `7.AnTX.` = lát An Taxi trong thư mục kế hoạch DÙNG CHUNG (khớp `file_glob` của 7 spec
        # atx_kehoach_*). Không lọc là kéo trùng với job SRVF/XDV và cả 4 khối khác.
        "chi_lay": r"^7\.AnTX\.",
    },
    # ── Xanh Vĩnh Phúc ─────────────────────────────────────────────────────────────────────
    # 1 file/tháng ('...M.2026.8.Baocaotonghop'), metadata bóc đúng month=8 nên KHÔNG cần
    # `ca_nam` như bên An Taxi. Nuôi spec `xvp_qtvh_ngay` (XVP_QTVH_D).
    {"company": "XANHVINHPHUC", "rt": "baocaoqtvhthang", "che_do": core.THANG,
     "ten": "QTVH Xanh Vĩnh Phúc theo tháng"},
    {
        "company": "KEHOACH", "rt": "baocaokehoachthang", "che_do": core.THANG,
        "ten": "Kế hoạch tháng Xanh Vĩnh Phúc",
        "chi_lay": r"^6\.XVP\.",
    },
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

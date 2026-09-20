#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QTVH AN TAXI — 3 lượt/ngày, xem khối MỐC GIỜ bên dưới.

Nuôi 5 màn `at1..at5` (Quản trị vận hành An Taxi). Khung chung + các bẫy của khung: xem
`cron_qtvh_core.py`. File này CHỈ khai báo nguồn.

TÁCH KHỎI `cron_antaxi_xvp_daily.py` NGÀY 29/08/2026 (user chốt). Bản gộp chạy đúng, nhưng gộp
hai khối vào một job thì cờ tắt `cron_<job>.disabled` tắt CẢ HAI — không dừng riêng được một khối
khi nguồn bên đó đang hỏng. Tách ra mỗi khối một artifact trạng thái, một log, một cờ tắt.

BA LƯỢT/NGÀY: 05:00 · 16:45 · 17:15 giờ VN (TEST) — cùng mốc với các job nguồn khác.
Lượt PROD chạy TRƯỚC TEST đúng 10 phút (04:50 · 16:35 · 17:05 VN). KHOẢNG CÁCH là bắt buộc,
chiều nào cũng được miễn khác phút: cùng phút thì hai lượt tranh khoá per-file
(servers/common/filelock.py), lượt sau bị 'skipped_lock' và MẤT HẲN một lượt nạp. Đánh đổi: test
hết vai 'chim báo bão' — nguồn hỏng nay prod vấp TRƯỚC, nên khi khối này im số thì đọc log PROD.
MỐC CRONTAB VIẾT THEO UTC (máy TZ=Etc/UTC, cron Ubuntu bỏ qua CRON_TZ khi TÍNH LỊCH — man 5
crontab, LIMITATIONS): TEST 09:45 · 10:15 · 22:00 UTC, PROD 09:35 · 10:05 · 21:50 UTC. Lượt sáng
(05:00 VN test / 04:50 VN prod) nằm ở 22:00 và 21:50 UTC HÔM TRƯỚC.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_antaxi_daily.py [--dry-run] [--env test|prod]

TRƯỚC KHI CÓ JOB NÀY (đo 29/08/2026) hai thư mục dưới đây kéo TAY hoàn toàn: bản trên đĩa dừng ở
28/08 16:34 VN trong khi bên agent đã có bản 29/08 16:36. Đó là lý do job này tồn tại.

VÌ SAO KHÔNG CÓ `baocaohqkdngay`: `cron_hqkdngay_daily.py` đã kéo (12 đơn vị dùng CHUNG report_type
đó) và đường `autofill` của nó gọi `spec_extract.run_for_path` nên các spec đọc thư mục đó vẫn
chạy theo. Khai lại ở đây là kéo trùng một file hai lần mỗi lượt.

VÌ SAO KHÔNG CÓ `baocaotaichinhrieng` / `baocaotuoino` / `baocaotaisancodinhcongcudungcu`: ba thư
mục đó là báo cáo TÀI CHÍNH THÁNG dùng chung cơ chế với MỌI đơn vị khác (đường derive_* cũ), không
riêng gì khối này. Tự động hoá chúng là một việc riêng, phải làm cho cả 12 đơn vị một lượt — làm
lẻ ở đây thì 11 đơn vị còn lại vẫn tay, mà lại tưởng là đã xong.
"""
import sys

import cron_qtvh_core as core

JOB = "antaxi_daily"
NHAN = "Nguồn QTVH An Taxi"
SCHEDULE_VN = "05:00 · 16:45 · 17:15 (prod sớm hơn 10')"   # 3 lượt/ngày. Ghi vào artifact cho
# agent giám sát khỏi hard-code mốc giờ — KHÔNG ai parse chuỗi này, chỉ hiển thị.

NGUON = [
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
        # atx_kehoach_*). Không lọc là kéo trùng với 3 job nguồn kia và cả 4 khối khác.
        "chi_lay": r"^7\.AnTX\.",
    },
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

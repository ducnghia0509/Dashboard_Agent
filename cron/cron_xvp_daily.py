#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QTVH XANH VĨNH PHÚC (XVP) — 3 lượt/ngày, xem khối MỐC GIỜ bên dưới.

Nuôi 6 màn `xt0..xt5` (Quản trị vận hành Xanh Taxi). Khung chung + các bẫy của khung: xem
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

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_xvp_daily.py [--dry-run] [--env test|prod]

TRƯỚC KHI CÓ JOB NÀY (đo 29/08/2026) `baocaoqtvhthang` kéo TAY hoàn toàn và bản trên đĩa còn là
bản 18/08 — TRỄ 11 NGÀY. Lượt chạy đầu tiên kéo dữ liệu từ 14/08 lên 23/08 (2.526 -> 3.239 dòng,
cuốc hoàn thành T8 124.006 -> 208.003). Đó là lý do job này tồn tại, không phải để chạy cho đủ bộ.

VÌ SAO KHÔNG CÓ `baocaohqkdngay` CỦA XVP: `cron_hqkdngay_daily.py` đã kéo (12 đơn vị dùng CHUNG
report_type đó) và đường `autofill` của nó gọi `spec_extract.run_for_path`, nên spec
`xvp_doanhthu_ngay` (XVP_DT_NGUON_D) cũng chạy theo — đã kiểm 29/08: file 29/08 trích ra 560 dòng,
DB có đúng 560. Khai lại ở đây là kéo trùng một file hai lần mỗi lượt.

VÌ SAO KHÔNG CÓ `baocaotaichinhrieng` / `baocaotuoino` / `baocaotaisancodinhcongcudungcu`: ba thư
mục đó là báo cáo TÀI CHÍNH THÁNG dùng chung cơ chế với MỌI đơn vị khác (đường derive_* cũ), không
riêng gì khối này. Tự động hoá chúng là một việc riêng, phải làm cho cả 12 đơn vị một lượt.

HTX XANH VĨNH PHÚC và HTX XANH TUYÊN QUANG là PHÁP NHÂN KHÁC, không thuộc job này: cả hai chỉ có
`baocaohqkdngay` (đã tự động) + tài chính tháng, KHÔNG có thư mục QTVH nào để kéo.
"""
import sys

import cron_qtvh_core as core

JOB = "xvp_daily"
NHAN = "Nguồn QTVH Xanh Vĩnh Phúc"
SCHEDULE_VN = "05:00 · 16:45 · 17:15 (prod sớm hơn 10')"   # 3 lượt/ngày. Ghi vào artifact cho
# agent giám sát khỏi hard-code mốc giờ — KHÔNG ai parse chuỗi này, chỉ hiển thị.

NGUON = [
    # 1 file/tháng ('...M.2026.8.Baocaotonghop'), metadata bóc đúng month=8 nên KHÔNG cần `ca_nam`
    # như bên An Taxi. MỖI SHEET LÀ MỘT THÁNG và tháng mới thêm sheet vào chính file cũ, nên kéo
    # lại là đè chính nó và vẫn giữ đủ các tháng cũ. Nuôi spec `xvp_qtvh_ngay` (XVP_QTVH_D).
    {"company": "XANHVINHPHUC", "rt": "baocaoqtvhthang", "che_do": core.THANG,
     "ten": "QTVH Xanh Vĩnh Phúc theo tháng"},
    {
        "company": "KEHOACH", "rt": "baocaokehoachthang", "che_do": core.THANG,
        "ten": "Kế hoạch tháng Xanh Vĩnh Phúc",
        # `6.XVP.` = lát Xanh Vĩnh Phúc trong thư mục kế hoạch DÙNG CHUNG (khớp `file_glob` của 7
        # spec xvp_kehoach_*). Không lọc là kéo trùng với 3 job nguồn kia và cả 4 khối khác.
        "chi_lay": r"^6\.XVP\.",
    },
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

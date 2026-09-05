#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QUẢN LÝ TÀI SẢN (QLTS) — 3 lượt/ngày, xem khối MỐC GIỜ bên dưới.

Nuôi 5 màn `ts1..ts5` (Quản lý tài sản). Khung chung + các bẫy của khung: xem
`cron_qtvh_core.py`. File này CHỈ khai báo nguồn.

VÌ SAO JOB NÀY TỒN TẠI (đo 29-30/08/2026)
-----------------------------------------
QLTS là khối DUY NHẤT có đủ 14 spec, 5 màn hoàn chỉnh, mà KHÔNG có dòng cron nào — mọi file
đều kéo tay. Hậu quả đo được lúc lập job:

  · toàn bộ 6 thư mục QLTS trên đĩa dừng ở bản 17/08, trong khi nguồn đã có bản 29/08
    (BCbaoduongxe / BCbaoduongmay / BCbaoduongxeDEMO) -> TRỄ 13 NGÀY, không ai biết;
  · trên kỳ đang xem (dataset 2026-08), màn `ts4` (Bảo hiểm/Đăng kiểm) và `ts5` (Nhiên liệu)
    TRỐNG HẲN — `thieuNguon: [QLTS_BH, QLTS_DK]` và `[QLTS_DAU]` — vì bản mới nhất của mấy
    nguồn đó còn nằm ở kỳ 2026-07.

HAI THƯ MỤC NGUỒN CHƯA PHÁT LẠI — KHAI SẴN, CÓ CHỦ Ý
----------------------------------------------------
`baocaotaisanqlts` (sổ tài sản, nuôi ts1/ts2 — 21.247 dòng) và `baocaodau` (nuôi ts5) hiện
KHÔNG có mặt trong danh mục metadata của agent .253: quét 745 file, không file nào tên chứa
`H5` + `Taisan`, cũng không có `BCDau`. Bản đang dùng là bản kéo tay ngày 17/08 nằm trên đĩa.

Vẫn khai hai thư mục đó ở đây vì hai lẽ: (1) ngày kế toán phát lại thì job tự kéo, không phải
nhớ sửa code; (2) mỗi lượt chạy sẽ để lại đúng một dòng "không có file nào" trong log — im lặng
mới là thứ đã để chuyện này kéo dài. KHÔNG được đọc dòng đó thành "job hỏng".

CHẾ ĐỘ `thang` CHO CẢ 6 THƯ MỤC
-------------------------------
Mọi nguồn QLTS đều là MỘT file cho MỘT kỳ, kỳ sau ra file mới (`...M.2026.8.` /
`...D.2026.8.` / `...Q.2026.7.`), nên kéo lại là đè chính nó — đúng định nghĩa `THANG`.
KHÔNG dùng `anh_chup_ky`: không nguồn nào ở đây phát nhiều bản chốt trong cùng một kỳ.

Kiểm kê là báo cáo QUÝ (`Q.2026.7`) nhưng metadata vẫn bóc `month=7`, nên nó rơi đúng vào cửa
sổ "tháng này + tháng trước" như mọi nguồn khác. Sang tháng 10 mà chưa có bản quý mới thì nó
tự rơi ra khỏi cửa sổ — đó là hành vi ĐÚNG, không phải mất dữ liệu: dòng cũ đã nằm trong DB.

BA LƯỢT/NGÀY: 05:00 · 16:45 · 17:15 giờ VN (TEST); PROD chạy TRƯỚC 10 phút (04:50 · 16:35 ·
17:05 VN). Khoảng cách giữa hai môi trường là BẮT BUỘC — cùng phút thì hai lượt tranh khoá
per-file và lượt sau mất hẳn. Crontab viết theo UTC (máy TZ=Etc/UTC): TEST 09:45 · 10:15 ·
22:00 UTC, PROD 09:35 · 10:05 · 21:50 UTC; lượt sáng nằm ở 22:00/21:50 UTC HÔM TRƯỚC.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_qlts_daily.py [--dry-run] [--env test|prod]
"""
import sys

import cron_qtvh_core as core

JOB = "qlts_daily"
NHAN = "Nguồn Quản lý tài sản (QLTS)"
SCHEDULE_VN = "05:00 · 16:45 · 17:15 (prod sớm hơn 10')"

NGUON = [
    # ── ts1 / ts2: sổ tài sản. CHƯA có trong metadata .253 (xem docstring) ──────────────
    {"company": "QLTS", "rt": "baocaotaisanqlts", "che_do": core.THANG,
     "ten": "Sổ tài sản (ts1, ts2)"},

    # ── ts2: kiểm kê quý ───────────────────────────────────────────────────────────────
    {"company": "QLTS", "rt": "baocaokiemke", "che_do": core.THANG,
     "ten": "Kiểm kê tài sản (ts2)"},

    # ── ts3: bảo dưỡng. MỘT thư mục, BA file khác nhau (máy / xe / xe DEMO) và 7 spec bóc
    # theo `file_glob` riêng từng file — nên KHÔNG cần `chi_lay`, cứ kéo cả ba.
    {"company": "QLTS", "rt": "baocaobaoduong", "che_do": core.THANG,
     "ten": "Bảo dưỡng máy / xe / xe DEMO (ts3)"},

    # ── ts4: bảo hiểm + đăng kiểm, hai thư mục riêng ───────────────────────────────────
    {"company": "QLTS", "rt": "baocaobaohiem", "che_do": core.THANG,
     "ten": "Theo dõi bảo hiểm (ts4)"},
    {"company": "QLTS", "rt": "baocaodangkiem", "che_do": core.THANG,
     "ten": "Theo dõi đăng kiểm (ts4)"},

    # ── ts5: nhiên liệu. CHƯA có trong metadata .253 (xem docstring) ───────────────────
    {"company": "QLTS", "rt": "baocaodau", "che_do": core.THANG,
     "ten": "Cấp phát nhiên liệu (ts5)"},
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

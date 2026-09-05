#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QTVH XƯỞNG DỊCH VỤ VINFAST (XDV) — 3 lượt/ngày, xem khối MỐC GIỜ bên dưới.

Nuôi 5 màn `xdv0..xdv4` (Quản trị vận hành xưởng dịch vụ). Khung chung + 4 cái bẫy: xem
`cron_qtvh_core.py`. File này CHỈ khai báo nguồn.

BA LƯỢT/NGÀY: 05:00 · 16:45 · 17:15 giờ VN (đổi 29/08/2026 theo yêu cầu user; trước đó một lượt 19:00 VN).
Lượt PROD chạy TRƯỚC TEST đúng 10 phút (04:50 · 16:35 · 17:05 VN; đổi 29/08/2026 theo yêu cầu
user, trước đó prod chạy SAU test 5 phút). KHOẢNG CÁCH là bắt buộc, chiều nào cũng được miễn khác
phút: cùng phút thì hai lượt tranh khoá per-file (servers/common/filelock.py), lượt sau bị
'skipped_lock' và MẤT HẲN một lượt nạp. Đánh đổi: test hết vai 'chim báo bão' — nguồn hỏng nay
prod vấp TRƯỚC, nên khi một khối im số thì đọc log PROD trước.
MỐC CRONTAB VIẾT THEO UTC (máy TZ=Etc/UTC, cron Ubuntu bỏ qua CRON_TZ khi TÍNH LỊCH — man 5
crontab, LIMITATIONS): TEST 09:45 · 10:15 · 22:00 UTC, PROD 09:35 · 10:05 · 21:50 UTC. Lượt sáng
(05:00 VN test / 04:50 VN prod) nằm ở 22:00 và 21:50 UTC HÔM TRƯỚC.

NHỊP NỘP THẬT (đo trên `available_metadata.json` 24/08/2026): doanh thu ngày nộp 16:04 giờ VN — hai
lượt chiều bắt được. Nhưng 5 NGUỒN KSCL TUẦN NỘP 18:02–18:31 VN, tức SAU CẢ HAI LƯỢT CHIỀU: bản của
ngày hôm nay chỉ vào DB ở LƯỢT 05:00 SÁNG HÔM SAU. Đó là lý do lượt sáng tồn tại — bỏ nó đi thì mỗi
bản KSCL trễ đúng một ngày, và đây cũng là điều PHẢI nói rõ khi ai đó hỏi "sao số KSCL chưa lên".

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
SCHEDULE_VN = "05:00 · 13:15 · 16:45 · 17:15 (prod sớm hơn)"  # 4 lượt/ngày (thêm lượt chiều 04/09/2026). Ghi vào artifact
# cho agent giám sát khỏi hard-code mốc giờ — KHÔNG ai parse chuỗi này, chỉ hiển thị.
# Lượt 05:00 sáng gánh các nguồn nộp sau giờ chiều (KSCL của XDV nộp 18:02-18:31 VN).
# Lượt prod chạy TRƯỚC test; xem khối chú thích trong crontab.
# LƯỢT CHIỀU 13:00 (prod) / 13:15 (test) thêm 04/09/2026: bộ nguồn tự động Cyber về VPS quanh
# trưa — bản 03/09 về lúc 12:14-12:19 VN — nên ba lượt cũ đều hụt, file phải đợi tới 16:45.

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
    # ── ảnh chụp TỪNG NGÀY (nguồn TỰ ĐỘNG Cyber -> ổ IT\TESTBAOCAOTUDONG\2.VINFAST_XDV): mỗi
    # file MỘT ngày rời nhau, GIỮ ĐỦ MỌI BẢN. Khai 04/09/2026; trước đó KHÔNG job nào kéo
    # `TEST_XDV`, file chỉ về VPS khi có người chạy tay nên chuỗi ngày đứt quãng mà không ai biết.
    #
    # `che_do` PHẢI là ANH_CHUP_KY: 9 file ngày mà để `luy_ke` thì bị coi là 9 bản chốt của một
    # tháng rồi XOÁ 8 cái — đúng cảnh báo đã ghi trong `_bay` của cả ba spec.
    #
    # Chỉ khai BA thư mục trong chín: ba cái này là "chỉ tiêu mới", nạp thẳng được và không đụng
    # report_type nào đang sống. Sáu thư mục còn lại của TEST_XDV cố ý ĐỂ NGOÀI (rà 03-04/09):
    #   · baocaotonghoplsc (DV42) và baocaolenhchuaxhd (DV44) — Cyber kéo dải 1 NGÀY nên ra dòng
    #     chảy, trong khi chỉ tiêu cần danh sách TREO luỹ kế; khai vào là số nhỏ hơn bản đang dùng
    #     hàng chục lần (đo 31/08: 578.769.141 đ / 11 xưởng vs 31,07 tỷ / 14 xưởng).
    #   · baocaodoanhthuxdvngay (DV62) — số tin cậy nhưng là chi tiết từng RO, còn ô đích là dòng
    #     theo mã chỉ tiêu B100/B110…; cần AGGREGATE mà engine chưa có mode đó.
    #   · cdpscongnotheoro (DV04) — rỗng 10/10 ngày. · baocaocongnotheohoadon (DV03) — sai bộ tài
    #     khoản (mapping ghi 13112/13117/13119/13812, dữ liệu thật ở 13131), mới có 1 ngày và rỗng.
    # (`baocaotaichinhrienghqkd` ĐÃ khai từ 04/09 — xem mục ngay dưới.)
    # Khai một thư mục chưa có spec thì mỗi lượt vẫn xin file về rồi nạp 0 dòng, đẻ log rác và một
    # dòng "nguồn chưa báo cáo" giả trong artifact giám sát.
    {
        "company": "TEST_XDV", "rt": "bangkehoadondv", "che_do": core.ANH_CHUP_KY,
        "ten": "Hoá đơn dịch vụ theo ngày (DV01 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        "company": "TEST_XDV", "rt": "bangkethungandv", "che_do": core.ANH_CHUP_KY,
        "ten": "Sổ quỹ thu ngân dịch vụ theo ngày (DV02 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        "company": "TEST_XDV", "rt": "bangkelenhsuachua", "che_do": core.ANH_CHUP_KY,
        "ten": "Bảng kê lệnh sửa chữa theo ngày (DV41 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        # BÁO CÁO LỢI NHUẬN KHỐI XDV theo ngày — nguồn LUỸ KẾ TỪ ĐẦU THÁNG, spec `xdv_pnl_ngay`
        # lấy HIỆU với bản ngày trước để ra số của riêng ngày (`tru_ngay_truoc`). Đây là nguồn mà
        # mapping XDV chỉ định cho doanh thu tổng/công việc/phụ tùng/khác và lợi nhuận gộp.
        #
        # PHẢI để ANH_CHUP_KY và phải giữ ĐỦ CHUỖI NGÀY trên đĩa: phép hiệu cần bản liền trước.
        # Bản nào về muộn hoặc bị bỏ thì ngày kế tiếp thành số GỘP của cả quãng — spec có cảnh báo
        # riêng cho ca đó, đừng bỏ qua nó trong log.
        "company": "TEST_XDV", "rt": "baocaotaichinhrienghqkd", "che_do": core.ANH_CHUP_KY,
        "ten": "Lợi nhuận khối XDV theo ngày (HQKD tự động, hiệu luỹ kế)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
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

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP NGUỒN QTVH SHOWROOM VINFAST (SRVF) — 3 lượt/ngày, xem khối MỐC GIỜ bên dưới.

Nuôi 5 màn `vhkd0..vhkd4` (Quản trị vận hành VHKD). Khung chung + 4 cái bẫy: xem
`cron_qtvh_core.py`. File này CHỈ khai báo nguồn.

BA LƯỢT/NGÀY: 05:00 · 16:45 · 17:15 giờ VN (đổi 29/08/2026 theo yêu cầu user; trước đó một lượt 11:00 VN).
Lượt PROD chạy TRƯỚC TEST đúng 10 phút (04:50 · 16:35 · 17:05 VN; đổi 29/08/2026 theo yêu cầu
user, trước đó prod chạy SAU test 5 phút). KHOẢNG CÁCH là bắt buộc, chiều nào cũng được miễn khác
phút: cùng phút thì hai lượt tranh khoá per-file (servers/common/filelock.py), lượt sau bị
'skipped_lock' và MẤT HẲN một lượt nạp. Đánh đổi: test hết vai 'chim báo bão' — nguồn hỏng nay
prod vấp TRƯỚC, nên khi một khối im số thì đọc log PROD trước.
MỐC CRONTAB VIẾT THEO UTC (máy TZ=Etc/UTC, cron Ubuntu bỏ qua CRON_TZ khi TÍNH LỊCH — man 5
crontab, LIMITATIONS): TEST 09:45 · 10:15 · 22:00 UTC, PROD 09:35 · 10:05 · 21:50 UTC. Lượt sáng
(05:00 VN test / 04:50 VN prod) nằm ở 22:00 và 21:50 UTC HÔM TRƯỚC.

NHỊP NỘP THẬT (đo 24/08/2026): nhóm "OO" (quản trị nội bộ VHKD) nộp 08:20–10:07 giờ VN, nên CẢ BA
lượt đều nằm sau mốc nộp — lượt 16:45 đã đủ, hai lượt còn lại là dự phòng cho file về muộn/gửi lại.
MỐC CRONTAB PHẢI VIẾT THEO UTC: máy đặt TZ=Etc/UTC và cron của Ubuntu bỏ qua CRON_TZ khi TÍNH
LỊCH (man 5 crontab, LIMITATIONS) -> 11:00 VN = 04:00 UTC.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_srvf_daily.py [--dry-run] [--env test|prod]

VÌ SAO KHÔNG CÓ `baocaohqkdngay` Ở ĐÂY: `cron_hqkdngay_daily.py` đã kéo nó cho cả 12 đơn vị (SRVF
là một trong đó). Khai lại là 2 job cùng ghi một `source_file`, và tuy có khoá per-file nên không
hỏng dữ liệu, lượt sau vẫn thấy vân tay do lượt trước tạo -> báo "chưa cập nhật" oan.

VÌ SAO KHÔNG CÓ `baocaotaichinhrieng` (BCTC tháng, cấp `DTHU_NHOM` = A100 của thẻ vhkd0): nó là
nguồn THÁNG, một kỳ một file, và `agent_cli autofill` trên file đó còn dẫn ra hơn chục report_type
tài chính dùng CHUNG TOÀN TẬP ĐOÀN (CĐKT, TSNV, PTHU, PTRA, THUE…). Nạp lại mỗi ngày là mỗi ngày
DELETE-then-INSERT cả khối đó — lỗi giữa đường là mất số của màn Tài chính, đổi lấy một con số chỉ
thay đổi mỗi tháng một lần. Job này chỉ LOG khi thấy kỳ mới ở nguồn (xem `canh_bao_bctc`), việc
nạp giữ nguyên bằng tay.
"""
import sys

import cron_qtvh_core as core

JOB = "srvf_daily"
NHAN = "Nguồn QTVH Showroom Vinfast (SRVF)"
SCHEDULE_VN = "05:00 · 13:15 · 16:45 · 17:15 (prod sớm hơn)"  # 4 lượt/ngày (thêm lượt chiều 04/09/2026). Ghi vào artifact
# cho agent giám sát khỏi hard-code mốc giờ — KHÔNG ai parse chuỗi này, chỉ hiển thị.
# Lượt 05:00 sáng gánh các nguồn nộp sau giờ chiều (KSCL của XDV nộp 18:02-18:31 VN).
# Lượt prod chạy TRƯỚC test; xem khối chú thích trong crontab.
# LƯỢT CHIỀU 13:00 (prod) / 13:15 (test) thêm 04/09/2026: bộ nguồn tự động Cyber về VPS quanh
# trưa — bản 03/09 về lúc 12:14-12:19 VN — nên ba lượt cũ đều hụt, file phải đợi tới 16:45.

NGUON = [
    # ── luỹ kế: mỗi kỳ NHIỀU bản chốt, chỉ giữ bản mới nhất rồi xoá rows bản cũ ─────────────
    # NGUỒN TAY 'Xuathoadon_' ĐÃ GỠ 03/09/2026 — `vhkd_kqkd` nghỉ hưu, `KDVH` nay lấy từ
    # TEST_SR/bangkehoadonbanxe (khai bên dưới). Giữ lại mục này thì mỗi lượt cron vẫn xin file về
    # rồi nạp 0 dòng (spec đã đổi đuôi .retired) — chỉ tổ đẻ log rác và một dòng "nguồn chưa báo
    # cáo" trong artifact giám sát.
    {
        # HỢP ĐỒNG KÝ MỚI bản THÁNG (KD60) — nuôi biểu đồ 3 và 4 của tab Báo cáo kinh doanh
        # (`vhkd1`). Cùng thư mục `SRVF/baocaokqkd` với sổ xuất hoá đơn tay vừa gỡ, cùng dạng tên
        # và cùng 3 kênh; `chi_lay` là thứ duy nhất tách hai loại file trong thư mục đó.
        # Mục này CHƯA nghỉ: bản tháng của hợp đồng ký mới vẫn là nguồn chính, nguồn ngày
        # (TEST_SR/baocaoxuathoadon) chỉ nối thêm phần sau 25/08.
        "company": "SRVF", "rt": "baocaokqkd", "che_do": core.LUY_KE,
        "ten": "Hợp đồng ký mới (VHKD_HDONG)",
        "chi_lay": r"Kymoi_",
        "slot": r"Kymoi_(B2B|B2C|GF)_T(\d+)",
        "ky_regex": r"Kymoi_(?:B2B|B2C|GF)_T(\d+)",
        "ngay_regex": r"\.M\.(20\d{2})\.(\d{1,2})\.\d+\.(\d{1,2})\.Kymoi",
    },
    # ── ảnh chụp TỪNG NGÀY: mỗi file một ngày rời nhau, GIỮ ĐỦ MỌI BẢN ──────────────────────
    # BA thư mục dưới đây là nguồn TỰ ĐỘNG (Cyber -> ổ IT\TESTBAOCAOTUDONG\1.VINFAST_SR). Trước
    # 03/09/2026 KHÔNG job nào kéo chúng: file chỉ về VPS khi có người chạy tay, nên các màn đọc
    # chúng đứng yên mà không ai biết. Hai trong ba (bán xe KD73, tồn kho KD36) nay là nguồn CHÍNH
    # của `KDVH` và `VHKD_TONVATLY` — hai nguồn tay tương ứng đã nghỉ hưu.
    #
    # `che_do` PHẢI là ANH_CHUP_KY, không phải LUY_KE: các file là những NGÀY RỜI NHAU chứ không
    # phải nhiều bản chốt của một kỳ. Để LUY_KE thì 9 file ngày bị coi là 9 bản chốt của tháng 8
    # rồi XOÁ 8 cái — đúng cảnh báo đã ghi sẵn trong `_bay` của hai spec tương ứng.
    {
        "company": "TEST_SR", "rt": "bangkehoadonbanxe", "che_do": core.ANH_CHUP_KY,
        "ten": "Xuất hoá đơn bán xe theo ngày (KD73 tự động)",
        # Tên file có HAI dạng tiền tố: 'B1.TC.TCKT.D.' (25-28/08) và 'B.1.TC.TCKT.D.' (từ 29/08)
        # -> regex cố ý bắt từ '.D.' trở đi, không bám tiền tố.
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        "company": "TEST_SR", "rt": "baocaonhapxuattonkhoxe", "che_do": core.ANH_CHUP_KY,
        "ten": "Tồn kho xe vật lý theo ngày (KD36 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        "company": "TEST_SR", "rt": "baocaoxuathoadon", "che_do": core.ANH_CHUP_KY,
        # TÊN THƯ MỤC NGƯỢC NGHĨA: đây KHÔNG phải báo cáo xuất hoá đơn mà là sổ theo dõi HỢP ĐỒNG
        # (KD60) — tiêu đề trong file là 'BẢNG KÊ ĐIỀU KIỆN HỢP ĐỒNG', có Ngày cọc / Hủy HĐ.
        # Sổ xuất hoá đơn là thư mục 'bangkehoadonbanxe' ngay trên.
        "ten": "Hợp đồng ký mới theo ngày (KD60 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        # KD66 — bảng tổng hợp bán xe theo showroom (đã cọc / TT đủ đã giao / TT đủ chưa giao).
        # Khai 04/09/2026 cùng lượt với KD25; cả hai là "chỉ tiêu mới", không đụng report_type nào
        # đang sống nên không cần ranh giới cutover như KD73.
        "company": "TEST_SR", "rt": "baocaotonghopbanxe", "che_do": core.ANH_CHUP_KY,
        "ten": "Tổng hợp bán xe theo ngày (KD66 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        # KD25 — nhập xuất tồn HỒ SƠ xe (giấy tờ/COC), khác KD36 là tồn kho xe VẬT LÝ.
        "company": "TEST_SR", "rt": "baocaonhapxuattonhosoxe", "che_do": core.ANH_CHUP_KY,
        "ten": "Tồn hồ sơ xe theo ngày (KD25 tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    # BA thư mục TEST_SR còn lại KHAI NỐT 05/09/2026. Cả ba đều ĐỔI NGUỒN CHỨ KHÔNG ĐỔI CÁCH TÍNH,
    # và cả ba đều để REPORT_TYPE RIÊNG chứ không ghi đè nguồn tay — nguồn nào cũng đang có một
    # khuyết tật ở phía FILE, nên nạp để đối chiếu song song trước, đổi sang sau khi khớp:
    #   · baocaocongnophaithungay -> `VHKD_PTHU_HD` (không phải `VHKD_PTHU_COC`): file THIẾU TRỌN
    #     tài khoản B2B 1316* ở cả 10/10 ngày, ghi đè là bốc hơi 57,41 tỷ công nợ B2B.
    #   · baocaotaichinhrienghqkd -> `VHKD_PNL_D` (không phải `HQKD_D`): cột doanh thu "Kỳ này"
    #     của file = 0 ở cả 3 ngày tháng 9, ghi đè là doanh thu Showroom về 0.
    #   · baocaotaichinhriengcdps -> `VHKD_CDPS_D`: NẠP ĐỂ SẴN, chưa nối màn nào (mapping có,
    #     thiết kế chưa khai ô đích).
    {
        # Cân đối PS công nợ theo hợp đồng (mapping VHKD dòng 17). Mỗi ngày một file, không luỹ kế.
        "company": "TEST_SR", "rt": "baocaocongnophaithungay", "che_do": core.ANH_CHUP_KY,
        "ten": "Công nợ theo hợp đồng theo ngày (tự động)",
        # Tên file dùng GẠCH DƯỚI trước 'Baocaocongnophaithu' chứ không phải dấu chấm như 5 thư mục
        # trên -> regex phải kết bằng '_', khớp '.' là không bắt được ngày nào.
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})_",
    },
    {
        # Báo cáo lợi nhuận toàn khối Showroom theo ngày (mapping VHKD dòng 11) — nguồn LUỸ KẾ
        # TỪ ĐẦU THÁNG, spec `vhkd_pnl_ngay` lấy HIỆU với bản ngày trước (`tru_ngay_truoc`), y hệt
        # cách đã làm cho XDV. PHẢI để ANH_CHUP_KY và phải giữ ĐỦ CHUỖI NGÀY trên đĩa: phép hiệu
        # cần bản liền trước, bản nào về muộn thì ngày kế tiếp thành số GỘP của cả quãng.
        "company": "TEST_SR", "rt": "baocaotaichinhrienghqkd", "che_do": core.ANH_CHUP_KY,
        "ten": "Lợi nhuận khối Showroom theo ngày (HQKD tự động, hiệu luỹ kế)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    {
        # Bảng CĐPS theo ngày (mapping VHKD dòng 3). Báo cáo DÒNG (phát sinh trong ngày).
        "company": "TEST_SR", "rt": "baocaotaichinhriengcdps", "che_do": core.ANH_CHUP_KY,
        "ten": "Cân đối phát sinh theo ngày (tự động)",
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})\.",
    },
    # NGUỒN TAY 'baocaotonkhoxevatly' ĐÃ GỠ 03/09/2026 — cùng lý do: `vhkd_tonkho_vatly` nghỉ hưu,
    # `VHKD_TONVATLY` nay lấy từ TEST_SR/baocaonhapxuattonkhoxe.
    {
        "company": "SRVF", "rt": "baocaokhoxeb2b", "che_do": core.LUY_KE, "ten": "Kho xe B2B",
        # Dạng năm.tháng.ngày ('2026.8.24.KHO XE'), 3 token — KHÔNG cùng dạng với 2 thư mục trên.
        # Bắt buộc khai: bản 17/08 vừa bị sửa lại muộn hơn bản 24/08 nên xếp theo giờ sửa là chọn
        # sai ảnh chụp (xem `core._xep_slot`).
        "ngay_regex": r"\.M\.(20\d{2})\.(\d{1,2})\.(\d{1,2})\.KHO",
    },
    {"company": "SRVF", "rt": "baocaokhoxeb2c", "che_do": core.LUY_KE, "ten": "Kho xe B2C"},
    {"company": "SRVF", "rt": "baocaokhoxegf", "che_do": core.LUY_KE, "ten": "Kho xe GF"},
    # BẮT BUỘC có `ky_regex`: 2 file 'BaocaoClaim_B2C_T11.25' / '_T12.25' (tháng 11-12 NĂM 2025)
    # bị bên quét gán month=7 theo token ngày tạo. Không khai thì chúng rơi vào slot của kỳ
    # 2026-07 và bị XOÁ như bản chốt cũ — mất dữ liệu claim 2025 thật (dry-run 24/08/2026 bắt).
    # Với regex này chúng ra kỳ 2026-11/2026-12 (năm suy từ token 20xx, '25' không phải năm) nên
    # nằm ngoài 2 kỳ đang kéo và job này không bao giờ chạm tới.
    {"company": "SRVF", "rt": "baocaoclaimb2b", "che_do": core.LUY_KE, "ten": "Claim B2B",
     "ky_regex": r"Claim_B2[BC]_T(\d+)",
     "ngay_regex": r"\.M\.(20\d{2})\.(\d{1,2})\.(\d{1,2})\.\s*Baocaoclaim"},
    {"company": "SRVF", "rt": "baocaoclaimb2c", "che_do": core.LUY_KE, "ten": "Claim B2C",
     "ky_regex": r"Claim_B2[BC]_T(\d+)",
     "ngay_regex": r"\.M\.(20\d{2})\.(\d{1,2})\.(\d{1,2})\.\s*Baocaoclaim"},
    {"company": "SRVF", "rt": "baocaonhapxeb2b", "che_do": core.LUY_KE, "ten": "Nhập xe B2B"},
    {"company": "SRVF", "rt": "baocaonhapxeb2c", "che_do": core.LUY_KE, "ten": "Nhập xe B2C"},
    {
        "company": "SRVF", "rt": "baocaocongnophaithu", "che_do": core.LUY_KE,
        # MỘT file cấp 3 report_type (VHKD_PTHU + _KENH + _COC) -> xoá bản cũ phải xoá TRỌN
        # source_file, không kèm report_type. Xem `core.xoa_ban_cu`.
        "ten": "Công nợ phải thu + COC",
        "ky_regex": r"Baocaocongnophaithu_T(\d+)",
        "ngay_regex": r"\.M\.(20\d{2})\.(\d{1,2})\.(\d{1,2})_Baocaocongnophaithu",
    },
    # ── tháng: 1 file/kỳ, kéo lại đè chính nó (idempotent theo source_file) ─────────────────
    {
        "company": "KEHOACH", "rt": "baocaokehoachthang", "che_do": core.THANG,
        "ten": "Kế hoạch tháng Showroom",
        # Thư mục KEHOACH gom kế hoạch của MỌI khối trong cùng report_type; `1.SR.` là phần Showroom
        # (khớp `file_glob` của 6 spec vhkd_kehoach_*). Không lọc là job này kéo cả kế hoạch Trạm
        # sạc/Xe tải/Xanh VP/An Taxi — ngoài phạm vi, và job XDV cũng kéo trùng.
        "chi_lay": r"^1\.SR\.",
    },
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

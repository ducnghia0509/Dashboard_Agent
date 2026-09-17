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
    {
        # KD23 — tồn kho xe theo HOÁ ĐƠN (khai 14/09/2026 theo bản mapping DASHBOARD 4 của VHKD).
        # Bộ ba tồn kho nay đủ mặt: KD36 vật lý · KD25 hồ sơ · KD23 hoá đơn. Ba tập KHÁC NHAU,
        # không cái nào thay được cái nào (14/09: hoá đơn 1.205 xe; 13/09: vật lý 1.602, hồ sơ
        # 1.844) — chính hiệu số vật lý − hoá đơn là chỉ tiêu khối B của màn vhkd4.
        #
        # `VHKD_TONHD` vẫn là REPORT_TYPE RIÊNG (không ghi đè `VHKD_KHOXE`), nhưng TỪ 15/09/2026
        # màn vhkd4 + CT40 ĐỌC NÓ TRƯỚC và chỉ lùi về 3 file tay ở kỳ nào chưa có nó — xem
        # `vhkd._ton_hoa_don`. Giữ hai report_type tách bạch chính là thứ cho phép đường lùi đó.
        "company": "TEST_SR", "rt": "baocaonhapxuatxetheosokhung", "che_do": core.ANH_CHUP_KY,
        "ten": "Tồn kho xe theo hoá đơn theo ngày (KD23 tự động)",
        # ĐÃ GỠ `chi_env: "test"` NGÀY 15/09/2026 — người dùng chốt chuyển hẳn màn vhkd4 sang
        # nguồn này, nên prod cũng kéo hằng ngày. Ba câu hỏi nghiệp vụ trong `_bay` của spec
        # được xử như sau thay vì chờ tiếp:
        #   · ký gửi/B2B_KD vắng mặt -> cứ nạp đúng cái file có, mã kho lạ lộ ra ở nhóm "Khác";
        #   · mất "chưa ghép khách" -> BE trả None + khai `thieuNguon`, FE ẩn thẻ (không in 0);
        #   · đổi cơ sở tuổi tồn -> in thẳng mốc đếm ra màn + ghi chú CT40, và loại xe demo/ký
        #     gửi khỏi CT40 để giữ nguyên phạm vi chỉ tiêu của các kỳ trước.
        # Kỳ T01–T08 vẫn xem bằng `VHKD_KHOXE` (3 file tay) qua đường lùi `vhkd._ton_hoa_don`:
        # file KD23 là sổ luỹ kế nhưng cột "Tồn cuối"/"Số ngày tồn" là ảnh chụp tại ngày chạy,
        # không xuất lại được ảnh chụp của quá khứ.
        # GẠCH DƯỚI trước tên báo cáo, giống `baocaocongnophaithungay` bên dưới chứ không giống 5
        # thư mục TEST_SR ở trên -> regex phải kết bằng '_'. Kết bằng '.' là không bắt được ngày
        # nào và mọi file rơi về slot rỗng.
        "ngay_regex": r"\.D\.(20\d{2})(\d{2})(\d{2})_",
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
    {
        # BẢNG CÂN ĐỐI KẾ TOÁN theo ngày (mapping VHKD dòng 10: "tất cả tài khoản trên BCĐKT, số dư
        # đầu kỳ, cuối kỳ"). Khai 07/09/2026 theo yêu cầu user — đây là thư mục TEST_SR CUỐI CÙNG
        # chưa được kéo, bị BỎ SÓT chứ không phải cố ý để ngoài như bên XDV: mapping đã đặt hàng từ
        # đầu, chỉ là chưa tới lượt làm. Trước đó trên đĩa đúng MỘT file (bản 05/09, kéo tay 06/09).
        #
        # NGUỒN KHÔNG CÂN — spec `vhkd_cdkt_ngay` nạp vào report_type RIÊNG `VHKD_CDKT_D` và CHƯA
        # NỐI MÀN NÀO. Đo 07/09/2026 trên bản 05/09: tổng tài sản 3.247,1 tỷ vs tổng nguồn vốn
        # −5.954,0 tỷ, lệch 9.201,1 tỷ ở cả ba cột. Xem `_bay` của spec trước khi gắn vào ô nào.
        #
        # ANH_CHUP_KY chứ không phải LUY_KE: mỗi ngày một file rời, để LUY_KE là các file ngày bị
        # coi là nhiều bản chốt của một tháng rồi xoá hết chỉ giữ một.
        "company": "TEST_SR", "rt": "baocaotaichinhriengcdkt", "che_do": core.ANH_CHUP_KY,
        "ten": "Bảng cân đối kế toán theo ngày (tự động)",
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
     # BẢN 9.15 CỦA T1 HỎNG BỐ CỤC (phát hiện 16/09/2026): nguồn chèn thêm 1 cột từ O trở đi trong
     # vùng DỮ LIỆU mà KHÔNG chèn ở dòng header (header vẫn 25 nhãn, dòng dữ liệu dài 26 ô) -> mọi
     # cột spec đọc đều lệch phải 1: 'Tình trạng hồ sơ' ra SỐ TIỀN, 'Tên DVCS' ra bộ hồ sơ. Nạp ra
     # 6 dòng rác nằm cạnh bản `…8.25.…T1` đúng (278 dòng / 10,0044 tỷ). Chỉ DUY NHẤT T1 lệch —
     # T2..T9 của cùng lô 9.15 đọc bình thường, nên không đụng spec, chỉ chặn đúng file này.
     # GỠ DÒNG NÀY khi nghiệp vụ phát hành lại T1 (tên mới, ngày khác 9.15) hoặc sửa cột trong
     # chính file 9.15 — để nguyên là bản T1 đúng về sau cũng bị chặn nếu họ giữ tên cũ.
     "bo_qua": r"\.9\.15\.\s*BaocaoClaim_B2B_T1\.",
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
        # ẢNH CHỤP SỐ DƯ: bản chốt muộn hơn ÍT DÒNG HƠN là bình thường (hợp đồng thu xong thì rụng
        # khỏi bảng), nên chốt 3 của `xoa_trung_ban_chot` được nới cho riêng nguồn này — xem
        # docstring hàm đó. Không khai thì mỗi kỳ hai bản chốt lại phải xoá tay: T9 đã dính đúng
        # vậy (bản 09.05 = 292 dòng/99,135 tỷ nằm cạnh bản 09.15 = 239 dòng/87,147 tỷ).
        "anh_chup_so_du": r"Baocaocongnophaithu",
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
    {
        "company": "SRVF", "rt": "baocaotiendoshowroomngay", "che_do": core.THANG,
        "ten": "Tiến độ giao xe theo ngày × showroom × dòng xe",
        # HAI file song song trong cùng thư mục, cùng bố cục, khác ý nghĩa:
        #   *_13_NGAY  -> KẾ HOẠCH  (spec `vhkd_kehoach_giaoxe_ngay`, VHKD_KH_GIAOXE_NGAY)
        #   *_30_NGAY  -> THỰC HIỆN (spec `vhkd_giaoxe_ngay`,         VHKD_TH_GIAOXE_NGAY)
        # `slot` BẮT BUỘC: thiếu nó thì hai file rơi vào CÙNG một slot (rt, kỳ) và bị coi là hai
        # bản chốt của nhau -> mỗi lượt giữ đúng một cái, cái kia bị xoá khỏi DB. Nhóm bắt được
        # ('13' / '30') tách chúng ra làm hai lát độc lập.
        "slot": r"_(\d+)_NGAY",
        # `ca_nam` BẮT BUỘC, cùng lý do như `ANTAXI/baocaoqtvhthang`: tên file KHÔNG có một chữ số
        # kỳ nào (`CHI_TIET_TIEN_DO_TUNG_SHOWROOM_TUNG_DONG_XE_13_NGAY.xlsx`) nên metadata trả
        # `month=null` và `thang_tu_ten_file` cũng chịu -> phép lọc `thang != month` loại file
        # khỏi MỌI lượt kéo, im lặng. Tên file cũng KHÔNG lăn theo tháng: sang tháng 10 vẫn đúng
        # cái tên đó, chỉ nội dung đổi -> kéo lại ở kỳ chính là đủ, và hai spec tự đọc kỳ từ dòng
        # tiêu đề trong sheet (`ky_thang_tu_o`) chứ không tin tên file.
        "ca_nam": True,
    },
]


if __name__ == "__main__":
    sys.exit(core.run(JOB, NHAN, NGUON, SCHEDULE_VN))

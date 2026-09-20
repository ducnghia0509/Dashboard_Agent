# -*- coding: utf-8 -*-
"""Deriver: BÁO CÁO NGÀY (HQKD theo từng ngày) — nguồn `received_reports/<FOLDER>/baocaohqkdngay/
B.<n>.<MÃ>.D.<YYYYMM>.<Tên>.xlsx` (ký tự **D** trong tên = Day; bản **M** là báo cáo tháng đã có
pipeline riêng, KHÔNG đụng tới).

PHẠM VI (chốt theo Mapping_Dashboard_QTTC.xlsx — cột "Đường link lấy dữ liệu ngày tạm thời trên
EXCEL", ngay bên phải cột "Map màn hình"): CHỈ đơn vị có ô này khác "ko có" mới lên được báo cáo
ngày, và CHỈ cụm chỉ tiêu P&L (doanh thu / giá vốn / chi phí / lợi nhuận). Mọi chỉ tiêu khác của
màn Công nợ · Tồn kho · Tài sản · Thuế · Dòng tiền ghi "ko có" -> KHÔNG dựng số ngày cho chúng.
Hiện có 10 đơn vị: SRVF, XANHVINHPHUC, HTXXANHTUYENQUANG, HTXXANHVINHPHUC, ANTAXI, ANKHACHSAN,
GLOBALAI, TRAMSAC, DUAN, HO.

report_type RIÊNG có hậu tố `_D` (HQKD_D / PNLT_D / CHIPHI_D) — KHÔNG ghi đè HQKD/PNLT/CHIPHI của
báo cáo tháng. LÝ DO BẮT BUỘC: dòng ngày nằm CÙNG dataset tháng (để bộ lọc "Khoảng ngày" của FE
vẫn resolve ra đúng dataset kỳ). Nếu dùng chung report_type thì chế độ Tháng — vốn KHÔNG gửi
from/to nên cộng hết mọi dòng trong dataset — sẽ cộng cả 31 dòng ngày lẫn dòng tháng = số gấp đôi.
Backend chọn `_D` hay bản tháng theo `grain` của request (xem app/metrics/repository.py::_rt).

Đơn vị & layout (verify trên file thật kỳ 2026-08):
  · SRVF  (`B.1.TC.TCKT.D.202608.BaocaoHQKD.xlsx`) — layout "srvf": mỗi ngày 1 sheet tên "{d}.{m}"
    (vd "1.8" = 01/08). Header dòng 1: Mục | Khối | Mã số | Chỉ tiêu | <Showroom> | % DT | …
    Mã A-series y hệt sheet tháng T{mm}BC (A100 doanh thu, A300 tổng CP, A310 giá vốn, A600 lợi
    nhuận, U302 LNST). KHÔNG có cột tổng (cột "CHI NHÁNH VINFAST HÀ NỘI" luôn = 0) -> chỉ ghi dòng
    theo cost center, tổng khối = Σ showroom (đúng nhánh cc_v của repository._per_file_resolved).
    ⚠ VỊ TRÍ CỘT LỆCH GIỮA CÁC SHEET (sheet "1.8" showroom đầu ở cột 5, "2.8"/"3.8" ở cột 6 vì
    thêm 1 cột "% DT") -> BẮT BUỘC dò cột theo TÊN header từng sheet, không hardcode index.
  · XANHVINHPHUC (`B.6.XVP.D.<YYYYMM>.Baocaohqkdngay.xlsx`) — layout "kqkd": sheet "01".."31" theo
    ngày. Header dòng 4: CHỈ TIÊU | MÃ SỐ | Tổng cộng | HO | Depot Phú Thọ | Depot Vĩnh Phúc |
    Depot Tuyên Quang. Có cột tổng -> ghi thêm dòng "trực tiếp" (không cost center) như bản tháng.
  · HTXXANHTUYENQUANG / HTXXANHVINHPHUC (`B.6.HTX_*.D.<YYYYMM>.Baocaotaichinhrieng.xlsx`) — cũng
    layout "kqkd" nhưng header ở dòng 6 và CHỈ 1 cột giá trị (không có cost center).
  · ANTAXI (`B.7.AAG.TCKT.**M**.<YYYYMM>.Baocaotaichinhrieng.xlsx`) — layout "antaxi": sheet
    "1".."31" theo ngày (+ sheet "TH" tổng hợp, bỏ qua). Header dòng 6: TT | CHỈ TIÊU | Tỉ lệ |
    Sơn Tây | Tỉ lệ | Thái Nguyên | Tỉ lệ | Tổng cộng | Xe thương quyền | … | Lũy kế tháng.
    ⚠ TÊN FILE ghi `.M.` (trùng hệt báo cáo THÁNG ở `baocaotaichinhrieng/`) — chỉ THƯ MỤC
    `baocaohqkdngay/` phân biệt được, xem `_period_of`/`_in_day_dir`.
    ⚠ Số mục La Mã ở cột TT riêng + lệch so với XVP -> bảng anchor riêng, xem `_ANTAXI_ANCHOR`.
  · ANKHACHSAN (`B.10.AAG.TCKT.D.<YYYYMM>.Baocaotaichinhrieng.xlsx`) — layout "anks", KHÁC HẲN
    4 layout trên: CHỈ 1 SHEET, mỗi NGÀY là 1 CỘT (không phải mỗi ngày 1 sheet). Header dòng 3:
    TT | Nội dung chi phí | ĐVT | Tổng cộng | 1 | 2 | … | 31 — số ngày nằm sẵn ở header, không
    suy từ tên sheet. Chỉ 1 cơ sở (Garden Sơn Tây), không cost center. Xem `_anks_all_days`.
  · GLOBALAI (`B.8.GA.TCKT.D.<YYYYMM><DD>.Baocaotaichinhrieng.xlsx`) — layout "tcode": mỗi ngày 1
    sheet tên "D1".."D31" (file LUỸ KẾ trong tháng — số sheet tăng dần theo ngày đã qua, KHÔNG
    zero-pad "D1" chứ không phải "D01"). Mã T100/T200/T300 ở cột A y hệt sheet THÁNG "HQKD HỢP
    NHẤT GA" (cột B = Chỉ tiêu, cột giá trị header ghi "D01".."D31" — CÓ zero-pad, khác tên sheet).
    KHÔNG cost center (GA 1 pháp nhân, không showroom/depot).
    ⚠ Tên file ghi thêm 2 số NGÀY snapshot ở cuối (".D.20260601." = kỳ 2026-06, KHÔNG PHẢI báo cáo
    của riêng ngày 01 — file là snapshot LUỸ KẾ cả tháng tính đến ngày lưu) — `_period_of` chỉ bắt
    6 số đầu sau ".D." nên vẫn ra đúng "2026-06" dù tên có thêm 2 số dư.
    Chỉ nạp 4 chỉ tiêu CÓ NGUỒN theo Mapping (DT thuần T101, Giá vốn T201, Tổng chi phí T200, LNTT
    T300) + breakdown chi phí T201-T204 (giá vốn/tài chính/bán hàng/QLDN, khớp nhãn TT200 mà bản
    THÁNG của GA đang dùng — xem `_TCODE_CANON`). KHÔNG suy diễn "Lợi nhuận sau thuế"
    từ T300: file NGÀY không tách dòng thuế TNDN, khác bản THÁNG (đang đọc sheet 'KQKD' TT200 hợp
    lệ, có dòng thuế riêng mã 51/52) — gán LNST=T300 như HT (T-series) sẽ SAI khi GA phát sinh
    thuế thật; để trống, tương tự các chỉ tiêu khác ghi "ko có" trong Mapping cho GA.
  · TRAMSAC (`B.3.TC.TCKT.D.<YYYYMMDD>.Baocaotaichinhrieng.xlsx`) — CÙNG layout "tcode" với GA
    (mã T100/T200/T300, sheet "D1".."D31" luỹ kế) nhưng KHÁC 1 ĐIỂM: quy ước PNLT LNTT/LNST.
    Verify DB 2026-06: PNLT của Trạm sạc CHỈ có đúng 1 dim1 'Lợi nhuận sau thuế' (không có 'Lợi
    nhuận trước thuế' riêng), giá trị BẰNG TUYỆT ĐỐI HQKD mã 1112 (-0,3844 = -0,3844) -> T300 ở
    Trạm sạc ĐÃ LÀ LNST (không như GA còn lăn tăn thuế TNDN chưa tách). Gán nhãn 'Lợi nhuận trước
    thuế' cho Trạm sạc như GA sẽ SAI TÊN cột trên bảng Cấu trúc Doanh thu/Chi phí (dim1 không khớp
    quy ước THÁNG). Xem tham số `profit_pnlt` của `_tcode_facts` — set qua `_UNITS["TRAMSAC"]`.
    ⚠ Tên file có NGÀY ĐẦY ĐỦ ".D.20260801." (8 số, khác GA chỉ dư 2 số) nhưng `_period_of` chỉ
    bắt 6 số đầu sau ".D." nên vẫn ra đúng "2026-08"; sheet "D1" KHÔNG zero-pad (khác header
    'D1' cũng không zero-pad, khác GA có header zero-pad 'D01' — không ảnh hưởng vì `_tcode_facts`
    dò val_j có fallback `ten_j + 1` khi không tìm thấy header dạng d\\d{2}).
  · DUAN (Khối Dự án, `B.4.TC.TCKT.D.<YYYYM>.BaocaoHQKD.xlsx`) — layout "duan": mỗi ngày 1 sheet
    tên "1".."31" (không zero-pad, giống "kqkd"). ⚠ THÁNG TRONG TÊN FILE KHÔNG ZERO-PAD ("
    .D.20268." = kỳ 2026-08, CHỈ 5 CHỮ SỐ — khác 6 chữ số của mọi đơn vị khác) → `_period_of`
    cần regex fallback riêng, xem bên dưới. Header CHIA 2 DÒNG khác nhau (dòng "STT|CHỈ TIÊU|
    DỰ ÁN" rồi dòng SAU MỚI liệt kê cột: "Tổng Dự án|HO Dự án|Cao Bằng|Tân Thịnh|Lạng Sơn|Yên
    Bình 3|Phú Quốc|Bình phước|Quang Sơn") — khác mọi layout khác (đều gộp nhãn cột + tên cost
    center CÙNG 1 dòng) nên KHÔNG dùng chung `_kqkd_scan`, có `_duan_facts` riêng. Cột "Tổng Dự
    án" = Σ 7 dự án (verify ngày 01/08: 392.421.525 + 239.351.852 = 631.773.377 = đúng cột Tổng)
    và "HO Dự án" luôn 0 ở dữ liệu đã verify → BỎ CẢ HAI, chỉ lấy 7 cột dự án (E:K) làm cost
    center, đúng theo Mapping ("E-K tương ứng cho các Costcenter"). Số La Mã (III/IV/V/VI/VIII/
    IX/X/XI/XII) nằm ở cột STT RIÊNG (giống "antaxi"), nhãn cột "CHỈ TIÊU" là text trần không số
    — neo theo nhãn CHUẨN HOÁ, đa số EXACT MATCH (không startswith) vì "Chi phí khác" (X.2, mã
    neo `cp_khac`) startswith sẽ trúng NHẦM dòng con "Chi phí khác tại dự án" (mục 1.7 của Giá
    vốn, đứng TRƯỚC trong sheet) nếu dùng prefix "chi phi khac" lỏng lẻo — chỉ 2 mã LNTT/LNST
    dùng startswith (nhãn có hậu tố "(EBT)"/"(EAT)" đổi được). Tổng chi phí = Giá vốn(IV) + Chi
    phí biến đổi(VI) + Chi phí cố định(VIII) + Chi phí tài chính(IX.2) + Chi phí khác(X.2) — ĐÚNG
    công thức "E25+E46+E74+E80+E83" của Mapping, cùng 5 nhóm CP y hệt ANTAXI (không có mục
    "phân bổ chung" cấp I riêng — đã nằm lồng trong "Chi phí khác"). Mã cost center Cao Bằng/
    Lạng Sơn/Phú Quốc/Quang Sơn lấy y hệt bản THÁNG (agent_cli._DA_PROJECT_CC: CB_DA/LS_DA/
    PQ_DA/QS_DA); "Tân Thịnh"→TT_DA và "Yên Bình"→YB_DA — ĐÃ VỀ XUÔI 28/08/2026, trước đó ngược
    (xem chú thích ở `agent_cli._DA_PROJECT_CC`). "Bình phước" CHƯA có trong master_data
    (giống Núi Pháo/Quảng Ngãi bản tháng) → mã tự đặt BINHPHUOC_DA, backfill cong_ty qua
    import_filled.

Neo dòng chỉ tiêu của layout "kqkd" theo TIỀN TỐ SỐ LA MÃ / số mục đã chuẩn hoá bỏ dấu, KHÔNG theo
"Mã số" và KHÔNG theo địa chỉ ô cứng (C17/C28/C100… như ghi chú trong file mapping): mã số bị TRÙNG
(HTX có 2 dòng mã "10": "1. Doanh thu bán hàng" và "3. Doanh thu thuần") và số thứ tự dòng LỆCH
giữa XVP (dòng 20) với HTX (dòng 22).

Tổng chi phí = Σ 6 cấu phần (IV giá vốn + VI biến đổi + VIII cố định + IX.2 tài chính + X.2 khác +
XII phân bổ chung) — verify khớp ĐÚNG dòng "18. Tổng chi phí" có sẵn trong file XVP ngày 01/08
(1.026.059.132 = 850.225.145 + 40.124.348 + 0 + 122.161.253 + 0 + 13.548.387). Lấy Σ cấu phần thay
vì dòng 18.1 để Σ lát CHIPHI_D luôn == tổng chi phí HQKD_D (file HTX không có dòng 18.1).

Chạy (dry-run, in tổng theo ngày, KHÔNG ghi):
  .venv/bin/python scripts/derive_hqkd_ngay.py <file.xlsx>
Ghi thật:
  .venv/bin/python scripts/derive_hqkd_ngay.py <file.xlsx> --write
"""
import argparse
import datetime
import json
import os
import re
import sys
import unicodedata

import openpyxl
import psycopg

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [p for p in (os.path.dirname(_HERE), _HERE) if p not in sys.path]
from servers.common import dataset_ky as _DSK  # noqa: E402
from servers.common import source_catalog as _SC  # noqa: E402

DB_URL = (os.environ.get("DATABASE_URL") or os.environ.get("TC_DATABASE_URL")
          or "postgresql://tc:tc@localhost:5433/tc_dashboard")

RT_HQKD, RT_PNLT, RT_CHIPHI, RT_DTHU = "HQKD_D", "PNLT_D", "CHIPHI_D", "DTHU_D"
# Cơ cấu GIÁ VỐN theo 7 khoản mục, riêng layout "duan" — xem `_DUAN_GV_CT` để biết vì sao phải là
# report_type riêng thay vì thêm dim1 vào RT_CHIPHI.
RT_DUAN_GV = "DUAN_GV_D"
# Cụm DÒNG CHẢY (cộng được theo khoảng), đối lại cụm SỐ DƯ đọc qua `snapshot_sum`. Dùng để hỏi
# "ngày này đã có P&L chưa" mà không đếm nhầm dòng số dư của cùng ngày — xem `_pl_quet_thu_muc`.
_RT_DONG_CHAY = frozenset((RT_HQKD, RT_PNLT, RT_CHIPHI, RT_DTHU))
# Doanh thu BÁN XE theo ngày × kênh (chỉ layout "srvf"). Tách khỏi RT_DTHU vì đó là doanh
# thu thuần toàn khối, còn cái này là cụm A200 chia B2C/B2B/GF — vế thực hiện của bảng
# điểm vhkd0. Tên khớp `KDVH` (bản THÁNG, nguồn BaocaoKQKD) + hậu tố _D theo quy ước ngày.
RT_KDVH = "KDVH_D"
# Lợi nhuận theo NGÀY x KENH (15/08/2026). Tach khoi PNLT_D: PNLT_D giu dong muc SHOWROOM
# (dim2 rong), con day la breakdown theo kenh — de chung mot report_type thi ai cong ca hai se
# dem doi. Verify sheet 13.8: A401 (B2C) -368.300.254 + A402 (B2B) -39.218.994 = -407.519.248
# = dung o A600 "LOI NHUAN SHOW ROOM".
RT_LN_KENH = "VHKD_LN_D"
# SỐ DƯ cuối NGÀY (stock) — KHÁC HẲN P&L ở trên (flow). Cộng dồn nhiều ngày là SAI, phía đọc
# phải lấy dòng của ngày mới nhất (xem `_snap_congno_facts`).
_RT_SODU = ("PTHU_D", "PTRA_D", "BS_D", "TS_D", "THUE_D", "HH_D", "TSNV_D")
REPORT_TYPES = (RT_HQKD, RT_PNLT, RT_CHIPHI, RT_DTHU, RT_KDVH, RT_LN_KENH) + _RT_SODU

# Mã chỉ tiêu 01_HQKD (khớp app/metrics/repository.py: HQKD_REVENUE/COST/PROFIT_AT).
MA_DT, MA_CP, MA_LNTT = "1000", "1047", "1112"

# Khai ở ĐÂY (không phải cạnh phần đọc nó) vì `_UNITS` bên dưới dùng làm khoá `layout_phu` — khối
# chú thích đầy đủ nằm ở layout "antaxi_bcqt".
_BCQT_SHEET = "BCQT PT."

_UNITS = {
    # `bo_tu_ngay` — CÙNG LÝ DO VỚI XDV bên dưới, xem chú thích ở đó. Showroom đã chuyển HQKD
    # ngày sang nguồn tự động `TEST_SR/baocaotaichinhrienghqkd` từ 01/09/2026 (bộ 4 spec
    # `vhkd_hqkd_ngay` / `vhkd_dthu_ngay` / `vhkd_pnlt_ngay` / `vhkd_chiphi_ngay`).
    # HÔM NAY MỐC NÀY KHÔNG ĐỔI GÌ, và đó chính là vấn đề: file tay đang mang tên
    # `D.202608` trong khi ruột nó CÓ sheet 01.09/02.09/03.09. Kỳ lấy từ TÊN FILE nên các sheet
    # tháng 9 bị bỏ qua — Showroom đang an toàn NHỜ MỘT CÁI TÊN SAI, không phải nhờ luật nào.
    # Kế toán Showroom đổi tên thành `D.202609` (đúng thứ IT đang đề nghị họ làm) là ngay hôm đó
    # nguồn tay bắt đầu ghi tháng 9 song song với nguồn tự động -> doanh thu/chi phí/LNST
    # Showroom GẤP ĐÔI, không có cảnh báo nào. Khai sẵn mốc để cái tên đúng không phá gì.
    "SRVF": {"layout": "srvf", "cong_ty": "TC", "khoi": "Khối KD Vinfast - Showroom",
             "bo_tu_ngay": "2026-09-01"},
    # BA KHỐI XANH — `sheet_sodu` KHAI THIẾU LÀ CÓ CHỦ Ý (18/09/2026). Khác Trạm sạc/An Taxi (một
    # workbook đủ 4 sheet số dư), ở đây mỗi đơn vị chỉ có một nửa DÙNG ĐƯỢC NGAY bằng hàm bóc của
    # bản THÁNG, nửa còn lại đúng khuôn nhưng khác mẫu nên phải viết thêm — khai bừa cả 4 khoá chỉ
    # đẻ ra dòng "bị loại" mỗi lượt chạy chứ không ra thêm số nào:
    #   · XVP     `CĐKT` = B01-DN, cột "Ngày cuối kỳ" -> `_snap_cdkt_facts`/`_snap_tsnv_facts` chạy
    #             thẳng. `CĐSPS` thì KHÔNG: header MỘT TẦNG ("Du no dau | Du co dau | Ps no | …")
    #             chứ không phải 2 tầng S06-DN của `_snap_2tang` -> Thuế/Hàng hoá chưa có nguồn.
    #   · HTX ×2  `CĐPS` đúng mẫu S06-DN 2 tầng -> chạy thẳng. `CĐKT` thì KHÔNG: mẫu B01-HTX
    #             (TT 71/2024), cột giá trị ghi "Kỳ này" chứ không "cuối kỳ", và MÃ SỐ 110/120/…
    #             mang nghĩa khác hẳn B01-DN nên `_TSNV_NHOM` map vào là sai nhóm chứ không phải
    #             thiếu dòng -> TS-NV/BS/TSCĐ chưa có nguồn.
    # CẢ BA đều KHÔNG có sổ tổng hợp công nợ trong file ngày -> không có `pthu`/`ptra`; công nợ
    # chi tiết theo đối tượng của 3 khối này nằm ở nguồn khác (`baocaotuoino/`), việc khác.
    "XANHVINHPHUC": {"layout": "kqkd", "cong_ty": "XVP", "khoi": "Khối KD Vận tải Taxi Xanh",
                     # `CĐKT` không khai "Từ ngày … Đến ngày …", chỉ một ô ngày — xem
                     # `_sodu_ngay_cua_wb` nhánh 2 và hai chốt chặn của nó.
                     "ky_sodu_o_ngay": True, "sodu_ho_file_rieng": True,
                     "sheet_sodu": {"cdkt": "CĐKT"}},
    # HAI HTX dùng mẫu CĐKT **B01-HTX** (TT 71/2024) chứ không phải B01-DN: cột giá trị ghi "Kỳ
    # này", tổng tài sản là mã 200 (không phải 270), TSCĐ là 150/151/152 (không phải 221/222/223)
    # — xem `_MAU_CDKT`. Khai `mau_cdkt` rồi thì `cdkt` dùng được, không phải viết hàm bóc mới.
    # Không có sổ công nợ chi tiết -> `_snap_sodu_facts` tự lấy TỔNG TK 131/331 từ CĐPS.
    # Kiểm chứng 19/09/2026: CĐKT ba ngày 15-16-17 CÂN TUYỆT ĐỐI (mã 200 = mã 500, lệch 0 đồng).
    # Từ 16/09/2026 họ file ảnh chụp `.D.<YYYYMMDD>.` của 2 HTX gánh LUÔN cả P&L: file tháng dừng
    # ở sheet "15" và không được xuất lại nữa, còn file ngày mọc thêm sheet "BC KQKD" (B02-HTX,
    # khai "Từ ngày 16/09 Đến ngày 16/09") + "HQKD" (cùng khuôn với sheet ngày của file tháng).
    # `pl_ho_file_rieng` bật `_pl_quet_thu_muc` vét những ngày file tháng còn thiếu.
    "HTXXANHTUYENQUANG": {"layout": "kqkd", "cong_ty": "HTX_XTQ", "khoi": "Khối KD Vận tải Taxi Xanh",
                          "sodu_ho_file_rieng": True, "mau_cdkt": "b01htx",
                          "pl_ho_file_rieng": {"sheet": "HQKD", "sheet_ky": "BC KQKD"},
                          "sheet_sodu": {"cdps": "CĐPS", "cdkt": "CĐKT"}},
    "HTXXANHVINHPHUC": {"layout": "kqkd", "cong_ty": "HTX_XVP", "khoi": "Khối KD Vận tải Taxi Xanh",
                        "sodu_ho_file_rieng": True, "mau_cdkt": "b01htx",
                        "pl_ho_file_rieng": {"sheet": "HQKD", "sheet_ky": "BC KQKD"},
                        "sheet_sodu": {"cdps": "CĐPS", "cdkt": "CĐKT"}},
    # HAI HỌ FILE trong CÙNG thư mục `ANTAXI/baocaohqkdngay/`, chọn theo NỘI DUNG chứ không theo
    # tên (tên `.M.`/`.D.` ngược hẳn nội dung — xem khối chú thích của layout "antaxi_bcqt"):
    #   · có sheet "BCQT PT."  -> `layout_phu` = "antaxi_bcqt" (họ `.D.<YYYYMMDD>`, NGUỒN MỚI)
    #   · không                -> "antaxi"                     (họ `.M.<YYYYMM>`,  nguồn cũ)
    # `bo_tu_ngay` CHỈ áp cho layout GỐC (xem `derive`), tức chỉ chặn họ `.M.` — nếu áp cho cả hai
    # thì nguồn mới vừa bật đã tự chặn chính mình. Mốc 2026-09-01 = ngày đầu tiên họ `.D.` có số.
    # ⚠ HAI NGUỒN KHÔNG ĐƯỢC CÙNG GHI MỘT NGÀY: khoá idempotent của chúng khác nhau (họ `.M.` theo
    # TÊN FILE, họ `.D.` theo KỲ) nên không xoá được của nhau -> chồng ngày là CỘNG ĐÔI, không phải
    # ghi đè. Mốc này là thứ duy nhất tách chúng ra.
    "ANTAXI": {"layout": "antaxi", "cong_ty": "AAG", "khoi": "Khối KD Dịch vụ An Taxi",
               "layout_phu": (_BCQT_SHEET, "antaxi_bcqt"), "bo_tu_ngay": "2026-09-01",
               # Cụm số dư: RUỘT giống hệt Trạm sạc (S31-DN / S06-DN / B01-DN), chỉ TÊN SHEET khác
               # -> khai tên ở đây, các hàm `_snap_*_facts` dùng chung không sửa dòng nào.
               "sheet_sodu": {"pthu": "131", "ptra": "331", "cdkt": "CDKT", "cdps": "CDPS"}},
    # AN KHÁCH SẠN — P&L ở file `.D.<YYYYMM>.` (layout "anks", mỗi ngày MỘT CỘT), số dư ở họ file
    # `.D.<YYYYMMDD>.` riêng, giống 3 khối Xanh nên cũng bật `sodu_ho_file_rieng` (18/09/2026).
    # KHÔNG khai `cdkt`: sheet đó là B01-DN nhưng cột giá trị ghi "Số cuối NĂM"/"Số đầu NĂM" thay vì
    # "cuối kỳ", `_snap_cdkt_facts`/`_snap_tsnv_facts` dò theo nhãn nên trả rỗng -> khai vào chỉ đẻ
    # dòng "bị loại" mỗi lượt. Ba sheet còn lại đúng mẫu chuẩn và chạy thẳng, lại có ĐỦ hai sổ công
    # nợ nên An KS ra được công nợ CHI TIẾT theo đối tượng (3 khối Xanh không có).
    "ANKHACHSAN": {"layout": "anks", "cong_ty": "AAG", "khoi": "Khối KD Dịch vụ An KS",
                   "sodu_ho_file_rieng": True,
                   "sheet_sodu": {"pthu": "131", "ptra": "331", "cdps": "CDPS"}},
    # pnlt_skip T101: GA bản THÁNG chạy extractor TT200 (không phải T-series) -> không có dim1
    # 'Doanh thu bán hàng'; xem `_tcode_facts`.
    "GLOBALAI": {"layout": "tcode", "cong_ty": "GA", "khoi": "Khối KD Công nghệ",
                 "pnlt_skip": ("T101",)},
    # CỐ Ý KHÔNG KHAI `bo_tu_ngay` CHO TRẠM SẠC (cân nhắc rồi bỏ, 06/09/2026) — dù từ bản
    # `.D.20260829.` file đã đổi layout (bỏ 31 sheet "D1".."D31", còn một sheet "BCHQKD" một cột
    # luỹ kế) y như Showroom/XDV lúc cutover. Khác một điểm quyết định: hai đơn vị kia có SPEC
    # KHÁC ghi số cho những ngày sau mốc, còn Trạm sạc thì KHÔNG CÓ nguồn nào thay. `bo_tu_ngay`
    # sẽ đổi trạng thái sang `da_chuyen_nguon` — một trạng thái LÀNH (xem `cron_status.py`) — tức
    # bật đèn xanh cho một luồng dữ liệu đã ĐỨT. Đèn đỏ `loi_nap` ở đây là ĐÚNG: file có, số
    # không có, cần người đi hỏi. Chỉ CÂU LÝ DO là sai hướng nên đã sửa ở `derive()`.
    #
    # VÌ SAO CHƯA DÙNG NGUỒN TỰ ĐỘNG (user chốt 06/09/2026: "không dùng nguồn đó"): BCHQKD của
    # bản tự động ghi ĐÚNG vế doanh thu (T101 tháng 8 = 1.420.972.130 đ, khớp tuyệt đối nguồn cũ;
    # Σ 5 dòng con = đúng T101) nhưng vế chi phí BỊ TRIỆT TIÊU DO LỆCH DẤU: "Chi phí giá vốn bán
    # trụ sạc" = −360.406.271 trong khi "Chi phí giá vốn tiền điện chi hộ" = +360.406.271, và
    # "Chi phí nhân viên" = −55.000 với "Phí ngân hàng" = +55.000 -> T200 = 0 và T300 = T100.
    # Nạp vế đó là LNST tháng 8 thành 1,421 tỷ thay vì 0,490 tỷ (gấp 2,9 lần).
    #
    # KỲ 2026-09: NGUỒN NGÀY ĐÃ ĐỨT, VÀ BỘ FILE MỚI KHÔNG THAY ĐƯỢC (chốt 06/09/2026).
    # Thư mục `baocaohqkdngay/` nay nhận mỗi ngày một file `.D.2026090X.` — nhìn tên thì tưởng là
    # báo cáo ngày mẫu mới, nhưng so từng tên sheet thì nó TRÙNG TUYỆT ĐỐI mẫu báo cáo THÁNG
    # `baocaotaichinhrieng/B.3.TC.TCKT.M.202607` (và cả `M.202508`): 11 sheet, cùng thứ tự. Tức
    # đây là mẫu THÁNG chụp lại theo ngày, số LUỸ KẾ từ đầu tháng — chi tiết + con số phóng đại
    # 3,26 lần nếu nạp sai đường: xem docstring `_tcode_day_sheets`.
    # Ngoài ra vế P&L còn TRỐNG: 'BCHQKD' của cả 5 file ghi 0 ở MỌI mã. Đã đối chiếu chéo 3 nguồn
    # trong CHÍNH file 05/09 để biết 0 đó thật hay do chưa điền:
    #   · bảng kê 'HĐ, chứng từ bán ra': TỔNG CỘNG = 0, không một dòng hoá đơn nào
    #   · 'BCĐPS' (01→05/09): mọi TK 511* có `phát sinh trong kỳ` = 0
    #   · 'Sổ nhật ký chung': 20 bút toán, toàn ghi tăng/giảm tài sản + khấu hao
    # => doanh thu tháng 9 = 0 là SỰ THẬT. Phát sinh duy nhất là khấu hao 1.695.395 đ/NGÀY đều
    # tăm tắp (6274 + 6424), tới 05/09 luỹ kế 8.476.975 đ — mà 'BCHQKD' vẫn ghi chi phí 0, tức
    # sheet P&L KHÔNG khớp sổ. Đó là việc phải hỏi bên sinh file, không sửa được ở đây.
    # Giữ nguyên `loi_nap`: file có, số ngày không có, cần người đi hỏi.
    #
    # MỞ LẠI 11/09/2026 — user chốt "cứ làm theo mapping" sau khi mapping
    # `Mapping_TramSac_daily.xlsx` được rà xong 3 cột BC tự động. `snapshot_sheet` bật nhánh
    # `_tcode_snap_per_day` cho bộ file "mỗi ngày một file, một sheet 'BCHQKD'". File
    # `.D.20260801.` vẫn còn 31 sheet "D1".."D31" -> vẫn đi nhánh cũ; nhánh này CHỈ chạy khi
    # không dò ra sheet ngày nào.
    #
    # ⚠ BỘ FILE NÀY ĐỔI BẢN CHẤT GIỮA CHỪNG, TÊN FILE KHÔNG ĐỔI MỘT KÝ TỰ:
    #   · tới 10/09/2026 — LUỸ KẾ, header ghi "Từ ngày 01/09 - Đến ngày 09/09"
    #   · 03:05 ngày 11/09 — kế toán xuất đè cả 10 file thành RIÊNG NGÀY, "Từ 08/09 - Đến 08/09"
    # Vì vậy `_tcode_snap_per_day` KHÔNG gán cứng chế độ nào: nó đọc dòng "Từ ngày .. Đến ngày .."
    # của từng file rồi tự quyết (xem `_snap_ky`). Kế toán đổi qua đổi lại cũng không phải sửa code.
    # ĐỪNG suy chế độ từ tên file hay từ một mốc ngày — lần đổi trên đã đi qua lặng lẽ đúng một
    # lượt và làm mọi con số thành rác (đọc luỹ kế trên file riêng ngày -> doanh thu tháng = 0).
    #
    # Đã kiểm chứng bộ file riêng ngày bằng 4 dấu hiệu độc lập (11/09/2026): header kỳ của 10/10
    # file; sổ nhật ký chung mỗi file chỉ chứa bút toán của đúng ngày đó; BCĐKT và BCĐPS có đầu kỳ
    # ngày N = cuối kỳ ngày N−1; và T100 GIẢM về 0 sau ngày 08 (luỹ kế không giảm được).
    # Σ 10 ngày = 282.317.354 = đúng số luỹ kế cả tháng của bản file hôm trước.
    "TRAMSAC": {"layout": "tcode", "cong_ty": "TC", "khoi": "Khối KD Trạm sạc Vgreen",
               "profit_pnlt": ("Lợi nhuận sau thuế",), "snapshot_sheet": "BCHQKD"},
    # SỐ DƯ THEO NGÀY (20/09/2026) — nuôi 4 màn Công nợ / Tồn kho / Tài sản / Thuế, vốn trống trơn
    # ở chế độ Ngày vì khối này là khối DUY NHẤT chưa khai `sheet_sodu` (đo trên DB 20/09: 0 dòng
    # `PTHU_D`/`PTRA_D`/`HH_D`/`TS_D`/`THUE_D`, trong khi 6 khối khác đều có).
    #
    # Nguồn là họ file ẢNH CHỤP `B.4.TC.TCKT.D.<YYYYMMDD>.Baocaotaichinhrieng.xlsx` nằm chung thư
    # mục với file P&L tháng, nên bật `sodu_ho_file_rieng` y như 3 khối Xanh và XDV.
    # TRƯỚC 20/09 HỌ FILE NÀY LÀ `.xlsb` nên mọi vòng quét bỏ qua từ khâu lọc đuôi (openpyxl từ
    # chối thẳng: "does not support binary format .xlsb"); kế toán đã đổi sang `.xlsx` cùng ngày.
    #
    # `ngay_sodu_tu_ten_file` bắt buộc ở đây: bản BCTC của kế toán Dự án BỎ TRỐNG ô kỳ — dòng 7 của
    # CĐKT nguyên văn "Tại ngày   tháng   năm ". Xem nhánh 4 của `_sodu_ngay_cua_wb`.
    #
    # `cdps` trỏ sheet `CDSPS` (mẫu này viết tắt khác, nội dung vẫn là S06-DN với hai tầng NỢ/CÓ
    # chuẩn) — nó nuôi HH_D (tồn kho) và THUE_D (thuế).
    "DUAN": {"layout": "duan", "cong_ty": "TC", "khoi": "Khối KD Dự án",
             "sodu_ho_file_rieng": True, "ngay_sodu_tu_ten_file": True,
             "sheet_sodu": {"pthu": "131", "ptra": "331", "cdkt": "CĐKT", "cdps": "CDSPS"}},
    "HO": {"layout": "ho_kqkd", "cong_ty": "TC", "khoi": "Khối hỗ trợ tập đoàn"},
    # XE TẢI HƯNG THỊNH (spec user 2026-08-06) — layout "ht", xem `_ht_facts`.
    "HUNGTHINH": {"layout": "ht", "cong_ty": "HT", "khoi": "Khối KD Xe tải"},
    # XƯỞNG DỊCH VỤ VINFAST (spec user 2026-08-06) — layout "xdv", xem `_xdv_facts`.
    # `bo_tu_ngay` = MỐC CUTOVER SANG NGUỒN TỰ ĐỘNG. Mapping XDV ('Báo cáo API_XDV' dòng 12) chốt:
    # `\\PHONGKETOANXUONGDICHVU\\BAOCAOHQKDNGAY` (file tay này) BỊ THAY bởi bản Cyber tự động
    # `B.2.TC.TCKT.D.2026mmdd.Baocaotaichinhrieng-HQKD` -> từ 01/09/2026 số ngày của khối XDV do bộ
    # bốn spec `xdv_hqkd_ngay` / `xdv_dthu_ngay` / `xdv_pnlt_ngay` / `xdv_chiphi_ngay` ghi.
    # Hai nguồn cùng ghi HQKD_D/DTHU_D/PNLT_D/CHIPHI_D + cùng khối + cùng 14 cost center, chồng
    # một ngày là GẤP ĐÔI (`_resolve_per_file` chỉ khử trùng TRONG một file, không khử xuyên file).
    # Mốc này phải TRÙNG `chi_nap_tu_ngay` của cả bốn spec — sửa một bên là hỏng.
    # Vì sao phải thay: bản tay `D.202609` sang tháng 9 chỉ có sheet 03.09 có số, 04.09 trở đi để
    # trắng, nên màn Tổng quan lọc khối XDV ngày 04/09 hiện 0 dù nguồn tự động đã đủ số.
    "XDV": {"layout": "xdv", "cong_ty": "TC", "khoi": "Khối KD Vinfast - XDV",
            "bo_tu_ngay": "2026-09-01"},
}

# Cost center theo TỪ KHOÁ trong tên cột (dò theo tên, không theo vị trí — xem docstring).
# Mã CC lấy y hệt bản tháng (agent_cli._SR_SHOWROOM_CC / raw_rows thật của XVP).
_CC_SRVF = [("uong bi", "UB_SR"), ("b2b", "B2B_SR"), ("oceanpark", "OCP_SR"), ("long bien", "LB_SR"),
            ("smart city", "SMC_SR"), ("ha long", "HL_SR"), ("cam pha", "CP_SR"),
            ("vinh phuc", "VP_SR"), ("son tay", "ST_SR"), ("xuan mai", "XM_SR")]
_CC_XVP = [("depot phu tho", "PT_DP"), ("depot vinh phuc", "VP_DP"), ("depot tuyen quang", "TQ_DP"),
           ("ho", "HO_XVP")]
# An Taxi: header ghi tên tỉnh trần ("Sơn Tây" / "Thái Nguyên"), mã CC lấy y hệt master_data
# (app/data/master_data.json: ST_AT 'Depot Sơn Tây', TN_AT 'Depot Thái Nguyên').
_CC_ANTAXI = [("son tay", "ST_AT"), ("thai nguyen", "TN_AT")]

# Showroom Uông Bí thuộc pháp nhân VFQN (SRVF là thư mục ĐA-pháp-nhân) — giống bản tháng.
_CC_CONGTY = {"UB_SR": "VFQN"}


def _nd(s):
    """Chuẩn hoá: bỏ dấu, thường hoá, gộp khoảng trắng (nguồn hay gõ sai dấu: 'Gíá vốn', 'LỢI NHUÂN')."""
    s = str(s or "").strip().lower().replace("đ", "d")
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s)


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _source_id(path):
    folder = (_SC.raw_company_from_path(path) or "").upper()
    return f"{folder}::{os.path.basename(path)}"


_DAY_DIR = "baocaohqkdngay"


def _in_day_dir(path):
    """File có nằm trong thư mục `baocaohqkdngay/` không."""
    return _nd(os.path.basename(os.path.dirname(os.path.abspath(path)))).replace(" ", "") == _DAY_DIR


def period_source_key(folder, period):
    """Khoá `raw_rows.source_file` của đơn vị nạp theo BỘ SNAPSHOT LUỸ KẾ (xem `_tcode_snap_per_day`).

    Ở đó một file KHÔNG phủ trọn tháng ngày như các layout khác — mỗi lượt nạp dựng lại cả kỳ từ
    cả bộ file, nên khoá idempotent phải là KỲ. HÀM CÔNG KHAI vì `cron_hqkdngay_daily.verify()`
    cũng phải hỏi DB bằng đúng khoá này: hỏi bằng khoá-theo-file sẽ ra KHONG_CO_DONG_NAO cho dữ
    liệu vừa nạp xong, tức bảng giám sát báo đỏ oan (gặp thật 11/09/2026, lượt chạy đầu sau khi
    bật snapshot cho Trạm sạc)."""
    return f"{folder}::{_DAY_DIR}::{period}"


def _period_of(file_name, in_day_dir=False):
    """'B.1.TC.TCKT.D.202608.BaocaoHQKD.xlsx' -> '2026-08'. None nếu không phải file kỳ .D.<YYYYMM>.

    An Taxi gửi báo cáo NGÀY nhưng đặt tên '.M.<YYYYMM>' Y HỆT báo cáo THÁNG
    ('B.7.AAG.TCKT.M.202608.Baocaotaichinhrieng.xlsx' — cùng tên với file trong
    `baocaotaichinhrieng/`) -> chỉ THƯ MỤC phân biệt được. Vì vậy '.M.' CHỈ được coi là kỳ báo
    cáo ngày khi `in_day_dir`; nếu nới lỏng theo tên file thì báo cáo THÁNG của An Taxi sẽ bị
    gate ngày bắt và pipeline tháng không bao giờ chạy."""
    m = re.search(r"\.D\.(\d{4})(\d{2})", file_name)
    if m is None:
        # DUAN: tháng KHÔNG zero-pad ('.D.20268.' = 2026-08, chỉ 5 chữ số) — bắt buộc neo dấu
        # '.' liền sau để không lẫn với 6-số của các đơn vị khác (regex trên đã tự fail cho
        # trường hợp đó nên không tranh chấp thứ tự thử).
        m = re.search(r"\.D\.(\d{4})(\d{1,2})\.", file_name)
    if m is None and in_day_dir:
        m = re.search(r"\.M\.(\d{4})(\d{2})", file_name)
    return f"{m.group(1)}-{int(m.group(2)):02d}" if m else None


def is_daily_report(path):
    """File này có phải BÁO CÁO NGÀY của đơn vị đã cấu hình không (dùng cho gate ở agent_cli).

    ĐÒI ĐÚNG THƯ MỤC `baocaohqkdngay/`, không chỉ đòi tên có `.D.<YYYYMM>` (siết 29/08/2026).
    An Taxi có `ANTAXI/baocaoqtvhngay/B.7.AAG.PKDVH.D.202608.Baocaotonghop.xlsx` — báo cáo QTVH
    theo ngày, do 6 spec `atx_ngay_*` bóc, KHÔNG liên quan HQKD. Chỉ khớp tên thì gate nhận nhầm,
    `derive()` chạy rồi trả 'không thấy sheet ngày nào khớp kỳ (layout antaxi)' và làm cả lượt
    autofill mang `ok:false` — trong khi các spec vẫn nạp đúng qua `run_for_path`. Hậu quả không
    phải mất số mà là MẤT TIN VÀO CẢNH BÁO: cron báo đỏ mỗi lượt cho một nguồn vẫn chạy tốt.
    Đã rà toàn bộ đĩa 29/08: mọi file `.D.` của 12 đơn vị HQKD đều nằm trong `baocaohqkdngay/`,
    nên siết ở đây không cắt mất nguồn thật nào.
    """
    folder = _source_id(path).split("::", 1)[0]
    if not _in_day_dir(path):
        return False
    return bool(_UNITS.get(folder)) and bool(_period_of(os.path.basename(path), True))


# ---------------------------------------------------------------------------------------------
# Layout "srvf" — mã A-series, mỗi ngày 1 sheet "{d}.{m}"
# ---------------------------------------------------------------------------------------------
# Cấu phần TRỰC TIẾP của A300 (bản tháng parse từ công thức A300; workbook data_only KHÔNG còn
# công thức nên dùng thẳng danh sách fallback y hệt agent_cli._chiphi_recs_srvf).
# Mã CON của A200 -> kênh bán. A211/A211A/A213 đều là B2B (xe khối B2B, xe khối B2B bản
# phụ, và B2B đã xuất hoá đơn) — gộp về một kênh cho khớp chiều phân tích của mapping
# ("Kênh B2B, B2C, GF"), mã gốc vẫn giữ ở dim3 để soát ngược từng dòng với file.
_SRVF_BANXE = [("A210", "B2C"), ("A211", "B2B"), ("A211A", "B2B"),
               ("A212", "GF"), ("A213", "B2B")]

# Ma CON cua A600 (LNTT) va U302 (LNST) -> kenh. Cung nguyen tac: chi doc dong CON, dong cha
# da nam o PNLT_D roi.
_SRVF_LN_KENH = [("A401", "B2C", "A600"), ("A402", "B2B", "A600"), ("A403", "GF", "A600"),
                 ("A407", "B2C", "U302"), ("A408", "B2B", "U302"), ("A409", "GF", "U302")]

_SRVF_CP_CODES = ["A310", "A320", "A325", "A330", "A340", "A350", "A360", "A500"]


def _srvf_day_sheets(wb, period):
    """[(sheet_name, 'YYYY-MM-DD')] — chỉ sheet tên '{ngày}.{tháng}' khớp THÁNG của kỳ."""
    y, mm = int(period[:4]), int(period[5:7])
    out = []
    for s in wb.sheetnames:
        m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})", s.strip())
        if m and int(m.group(2)) == mm and 1 <= int(m.group(1)) <= 31:
            out.append((s, f"{y:04d}-{mm:02d}-{int(m.group(1)):02d}"))
    return sorted(out, key=lambda x: x[1])


def _srvf_facts(rows):
    """rows -> [(cost_center, report_type, dim1, dim3, value_VND)] cho 1 ngày. [] nếu sai layout."""
    hdr = next((r for r in rows[:8] if any(_nd(c) == "ma so" for c in r if c is not None)), None)
    if hdr is None:
        return []
    ma_j = next(j for j, c in enumerate(hdr) if _nd(c) == "ma so")
    ten_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("chi tieu")), ma_j + 1)
    cols = [(cc, j) for kw, cc in _CC_SRVF
            for j in [next((j for j, c in enumerate(hdr) if isinstance(c, str) and kw in _nd(c)), None)]
            if j is not None]
    if len(cols) < 8:      # thiếu showroom = layout lạ -> không đoán
        return []
    by_code = {}
    for r in rows:
        c = str(r[ma_j]).strip().upper() if ma_j < len(r) and r[ma_j] not in (None, "") else ""
        if c and c not in by_code:
            by_code[c] = r

    def val(code, j):
        r = by_code.get(code)
        return _num(r[j]) if r is not None and j < len(r) else None

    from extract_chiphi import _nhom_cp
    facts = []
    for cc, j in cols:
        for code, rt, dim1 in (("A100", RT_HQKD, MA_DT), ("A300", RT_HQKD, MA_CP),
                               ("A600", RT_HQKD, MA_LNTT), ("A310", RT_PNLT, "Giá vốn hàng bán"),
                               ("A100", RT_PNLT, "Doanh thu HH, DV"),
                               ("A600", RT_PNLT, "Lợi nhuận trước thuế"),
                               ("A100", RT_DTHU, "Doanh thu thuần")):
            v = val(code, j)
            if v:
                facts.append((cc, rt, dim1, dim1, v))
        lnst = val("U302", j)
        lnst = lnst if lnst else val("A600", j)
        if lnst:
            facts.append((cc, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lnst))
        # ── Doanh thu BÁN XE theo KÊNH (mapping VHKD dòng 37: "Số liệu BC ngày: Folder
        # BAOCAOHQKDNGAY - Sheet 01,02.. tương ứng ngày, tương ứng các cột costcenter").
        # CHỈ lấy các mã CON, KHÔNG lấy A200: A200 = A210+A211+A211A+A212+A213 (verify sheet 13.8:
        # 7.352.955.365 + 387.727.273 + 1.697.181.819 = 9.437.864.457 = đúng ô A200). Lấy cả hai
        # là gấp đôi doanh thu bán xe của mọi ngày.
        for code, kenh in _SRVF_BANXE:
            v = val(code, j)
            if v:
                facts.append((cc, RT_KDVH, "A200", code, v, kenh))
        # Loi nhuan truoc/sau thue theo KENH. dim1 = ma CHA (A600 / U302) de tang chung mot
        # truc voi PNLT_D, dim3 = ma con de soat nguoc tung dong voi file.
        for code, kenh, cha in _SRVF_LN_KENH:
            v = val(code, j)
            if v:
                facts.append((cc, RT_LN_KENH, cha, code, v, kenh))

        for code in _SRVF_CP_CODES:
            v = val(code, j)
            if not v:
                continue
            r = by_code[code]
            ten = str(r[ten_j]).strip() if ten_j < len(r) and r[ten_j] not in (None, "") else code
            facts.append((cc, RT_CHIPHI, _nhom_cp(ten, code), ten, v))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "kqkd" — XVP / HTX, mỗi ngày 1 sheet "01".."31"
# ---------------------------------------------------------------------------------------------
# (khoá, tiền tố nhãn đã chuẩn hoá) — neo theo NHÃN, xem docstring đầu file.
# ⚠ "lntt"/"lnst" dùng CONTAINS (tuple 1 phần tử), KHÔNG theo tiền tố La Mã như các dòng khác:
# verify file thật 2026-08 (XVP/HTX_XTQ/HTX_XVP) — cột "Mã số" đúng thứ tự (…11,12,13,14,15…) nhưng
# CỘT NHÃN kế toán ghi SAI/TRÙNG số La Mã từ dòng 13 trở đi (vd XVP: mã 13 "Lợi nhuận trước thuế
# TNDN (EBT)" lại ghi "XI." — trùng roman của chính mã 11 "LN TRỰC TIẾP TRƯỚC THUẾ TNDN" đứng trước;
# mã 15 "Lợi nhuận sau thuế TNDN (EAT)" ghi "XIII." thay vì "XV"; HTX ghi "XII."/"XIII." tuỳ đơn vị)
# -> neo prefix "xiii."/"xv." cũ KHÔNG BAO GIỜ khớp, HQKD/PNLT mất trắng 1112 + "Lợi nhuận trước/sau
# thuế" ở cả 3 đơn vị (phát hiện 2026-08-07, dữ liệu TEST xác nhận thiếu qua raw_rows). "tndn" bắt
# buộc trong "lntt" để KHÔNG trúng nhầm dòng "...TRỰC TIẾP TRƯỚC THUẾ TNDN" (mã 11, đứng TRƯỚC dòng
# đúng trong sheet) — 2 dòng đều chứa "loi nhuan" + "truoc thue" nhưng chỉ dòng EBT có liền cụm
# "loi nhuan truoc thue tndn" (chữ "truc tiep" chen giữa ở dòng kia phá vỡ tính liền mạch).
_KQKD_ANCHOR = {
    "dt_thuan": "3. doanh thu thuan",
    "gia_von": "iv. gia von hang ban",
    "ln_gop": "v. loi nhuan gop",
    "cp_bien_doi": "vi. chi phi bien doi",
    "cp_co_dinh": "viii. chi phi co dinh",
    "cp_tai_chinh": "2. chi phi tai chinh",
    "cp_khac": "2. chi phi khac",
    "pb_chung": "xii. phan bo chi phi chung",
    "lntt": ("loi nhuan truoc thue tndn",),
    "lnst": ("loi nhuan sau thue",),
}
# Cấu phần tổng chi phí -> (nhóm CP chuẩn mực, nhãn hiển thị). Nhóm khớp bản THÁNG của cùng đơn vị
# (raw_rows CHIPHI của XVP chỉ có 4 nhóm: Giá vốn / Chi phí tài chính / Chi phí bán hàng / Chi phí
# QLDN) để biểu đồ cơ cấu chi phí đọc được như nhau ở cả 2 chế độ; nhãn gốc giữ ở dim3.
_KQKD_CP = [
    ("gia_von", "Giá vốn hàng bán", "Giá vốn hàng bán"),
    ("cp_bien_doi", "Chi phí bán hàng", "Chi phí biến đổi"),
    ("cp_co_dinh", "Chi phí QLDN", "Chi phí cố định"),
    ("cp_tai_chinh", "Chi phí tài chính", "Chi phí tài chính"),
    ("cp_khac", "Chi phí khác", "Chi phí khác"),
    ("pb_chung", "Chi phí QLDN", "Phân bổ chi phí chung"),
]


def _kqkd_day_sheets(wb, period):
    y, mm = int(period[:4]), int(period[5:7])
    out = []
    for s in wb.sheetnames:
        t = s.strip()
        if t.isdigit() and 1 <= int(t) <= 31:
            out.append((s, f"{y:04d}-{mm:02d}-{int(t):02d}"))
    return sorted(out, key=lambda x: x[1])


def _kqkd_scan(rows, ccs, anchors):
    """Dò header + cột giá trị + neo dòng chỉ tiêu cho layout dạng KQKD (dùng cho cả 'kqkd' và
    'antaxi'). -> (cols, anchored, val) ; cols=[] nếu sai layout."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None:
        return [], {}, None
    hdr = rows[hdr_i]
    ten_j = next(j for j, c in enumerate(hdr) if _nd(c) == "chi tieu")
    # Cột giá trị: ƯU TIÊN các cột cost center (Depot/HO/tỉnh). CỐ Ý BỎ cột "Tổng cộng" khi đã có
    # cột cost center — Σ cost center khớp ĐÚNG cột tổng (verify XVP 01/08: 283.756.022 +
    # 480.309.773 + 140.030.764 = 904.096.560; An Taxi 01/08 DT thuần: 76.225.625 + 34.226.640 =
    # 110.452.265), nên ghi thêm dòng tổng chỉ tạo nguy cơ đếm đôi ở các truy vấn KHÔNG đi qua
    # repository._per_file_resolved (daily_series/sum_by_dim1 — biểu đồ xu hướng ngày).
    # Đơn vị 1 cột (HTX, không có cost center) -> dùng cột tổng / cột ngay sau "MÃ SỐ".
    cols = [(cc, j) for j, c in enumerate(hdr) if j != ten_j
            for cc in [next((cc for kw, cc in ccs if _nd(c) == kw), None)] if cc]

    anchored = {}
    for r in rows[hdr_i + 1:]:
        if not r or ten_j >= len(r):
            continue
        n = _nd(r[ten_j])
        for key, pref in anchors.items():
            if key in anchored:
                continue
            # pref là tuple 1 phần tử -> khớp CONTAINS (bất kể tiền tố La Mã đứng trước, xem
            # comment `_KQKD_ANCHOR` lntt/lnst); chuỗi thường -> khớp STARTSWITH như cũ.
            hit = (pref[0] in n) if isinstance(pref, tuple) else n.startswith(pref)
            if hit:
                anchored[key] = r

    def val(key, j):
        r = anchored.get(key)
        return _num(r[j]) if r is not None and j < len(r) else None

    # Đơn vị 1 cột (HTX, không cost center) -> cột "Tổng cộng", không có thì cột số ĐẦU TIÊN bên
    # phải "MÃ SỐ". KHÔNG lấy cứng `ma_j + 1` (19/09/2026): sheet "HQKD" trong file NGÀY của HTX
    # chèn thêm cột "Thuyết minh" giữa "MÃ SỐ" và cột số, còn sheet ngày trong file THÁNG thì
    # không — lấy cứng là trúng ô rỗng và CẢ SHEET ra 0 dòng mà không một câu lỗi nào (9/9 neo
    # vẫn khớp, chỉ `val` trả None). Dò bằng chính dòng đã neo nên tự đúng cho cả hai khuôn; cột
    # "% DT" đứng sau cột số nên không bao giờ được chọn trước.
    if not cols:
        tong_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("tong cong")), None)
        if tong_j is None:
            ma_j = next((j for j, c in enumerate(hdr) if _nd(c) == "ma so"), ten_j)
            moc = anchored.get("dt_thuan")
            tong_j = next((j for j in range(ma_j + 1, len(hdr))
                           if moc is not None and j < len(moc) and _num(moc[j])), ma_j + 1)
        cols = [(None, tong_j)]

    return cols, anchored, val


def _kqkd_facts(rows):
    """rows -> [(cost_center|None, report_type, dim1, dim3, value_VND)]. [] nếu sai layout."""
    cols, anchored, val = _kqkd_scan(rows, _CC_XVP, _KQKD_ANCHOR)
    if not cols or "dt_thuan" not in anchored:
        return []

    facts = []
    for cc, j in cols:
        dt = val("dt_thuan", j)
        if dt:
            facts.append((cc, RT_HQKD, MA_DT, MA_DT, dt))
            facts.append((cc, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
            facts.append((cc, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
        tong_cp = 0.0
        for key, nhom, ten in _KQKD_CP:
            v = val(key, j)
            if not v:
                continue
            tong_cp += v
            facts.append((cc, RT_CHIPHI, nhom, ten, v))
        if tong_cp:
            facts.append((cc, RT_HQKD, MA_CP, MA_CP, tong_cp))
        for key, rt, dim1 in (("lntt", RT_HQKD, MA_LNTT), ("gia_von", RT_PNLT, "Giá vốn hàng bán"),
                              ("ln_gop", RT_PNLT, "Lợi nhuận gộp"),
                              ("lntt", RT_PNLT, "Lợi nhuận trước thuế"),
                              ("lnst", RT_PNLT, "Lợi nhuận sau thuế")):
            v = val(key, j)
            if v:
                facts.append((cc, rt, dim1, dim1, v))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "antaxi" — An Taxi (ANTAXI/baocaohqkdngay), mỗi ngày 1 sheet "1".."31" + sheet "TH"
# ---------------------------------------------------------------------------------------------
# KHÁC layout "kqkd" ở 2 điểm (verify file thật B.7.AAG.TCKT.M.202608, sheet 1-31, header dòng 6):
#  1. Số mục La Mã nằm ở cột "TT" RIÊNG, KHÔNG dính vào nhãn: dòng 33 = TT "III" + CHỈ TIÊU
#     "DOANH THU THUẦN" (XVP ghi "3. Doanh thu thuần" trong CÙNG 1 ô) -> neo theo NHÃN TRẦN.
#  2. Số La Mã của An Taxi lệch hẳn XVP (LNTT = XI ở An Taxi vs XIII ở XVP; LNST = XIII vs XV;
#     An Taxi KHÔNG có mục "phân bổ chi phí chung") -> phải có bảng anchor riêng, không dùng chung.
# Đã verify nhãn nào cũng khớp DUY NHẤT 1 dòng trên cả 31 sheet (các dòng con "Giá vốn bán xe…",
# "Chi phí hoạt động khác", "LỢI NHUẬN TRƯỚC KHẤU HAO…(EBITDA)" đều KHÔNG startswith nhãn neo).
_ANTAXI_ANCHOR = {
    "dt_gross": "doanh thu ban hang",              # I  (mã 100) — DT GỘP
    "giam_tru": "cac khoan giam tru doanh thu",    # II (mã 110)
    "dt_thuan": "doanh thu thuan",                 # III (mã 120)
    "gia_von": "gia von hang ban",                 # IV (mã 130)
    "ln_gop": "lai gop",                           # V  (mã 140) — An Taxi ghi "LÃI GỘP"
    "cp_bien_doi": "chi phi bien doi",             # VI (mã 150)
    "cp_co_dinh": "chi phi co dinh",               # VIII (mã 170)
    "dt_tai_chinh": "doanh thu tai chinh",         # IX.1
    "cp_tai_chinh": "chi phi tai chinh",           # IX.2 (mã 182)
    "tn_khac": "thu nhap khac",                    # X.1
    "cp_khac": "chi phi khac",                     # X.2 (mã 192)
    "lntt": "loi nhuan truoc thue",                # XI (mã 200)
    "lnst": "loi nhuan sau thue",                  # XIII (mã 220)
}
# Tổng chi phí = IV + VI + VIII + IX.2 + X.2 (đúng công thức cột "Công thức" của mapping An Taxi;
# file KHÔNG in dòng tổng chi phí nào để đối chiếu). Verify vòng kín theo LNTT có sẵn trong file —
# 01/08 Sơn Tây: 76.225.625 (DTT) − 93.342.155 (Σ 5 cấu phần) + 2.965.749 (thu nhập khác) + 0
# (DT tài chính) = −14.150.781 = ĐÚNG dòng "XI. LỢI NHUÂN TRƯỚC THUẾ TNDN (EBT)".
# dim1 (nhóm CP) lấy Y HỆT bản THÁNG của CHÍNH An Taxi — raw_rows CHIPHI T6 có đúng 5 nhóm:
# Giá vốn hàng bán / Chi phí biến đổi / Chi phí cố định / Chi phí tài chính / Chi phí khác
# (KHÁC XVP dùng 'Chi phí bán hàng'/'Chi phí QLDN') để biểu đồ cơ cấu chi phí đọc được như nhau
# ở cả 2 chế độ Ngày/Tháng.
_ANTAXI_CP = [
    ("gia_von", "Giá vốn hàng bán", "Giá vốn hàng bán"),
    ("cp_bien_doi", "Chi phí biến đổi", "Chi phí biến đổi"),
    ("cp_co_dinh", "Chi phí cố định", "Chi phí cố định"),
    ("cp_tai_chinh", "Chi phí tài chính", "Chi phí tài chính"),
    ("cp_khac", "Chi phí khác", "Chi phí khác"),
]


def _antaxi_emit(cc, get):
    """(cost_center, get(key) -> số|None) -> [fact]. BỘ FACT CHUNG cho CẢ HAI họ file An Taxi.

    Tách khỏi `_antaxi_facts` ngày 16/09/2026 khi mở thêm layout "antaxi_bcqt". Hai họ file cùng
    nuôi MỘT đơn vị và nằm hai bên một mốc cutover, nên bắt buộc sinh Y HỆT bộ dim1 — lệch một
    nhãn là ngày trước mốc và ngày sau mốc rơi vào hai nhóm khác nhau trên cùng một biểu đồ cơ
    cấu, mà không có cảnh báo nào (số vẫn có, tổng vẫn đúng, chỉ cột bị tách đôi).
    """
    facts = []
    dt = get("dt_thuan")
    if dt:
        facts.append((cc, RT_HQKD, MA_DT, MA_DT, dt))
        facts.append((cc, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
    tong_cp = 0.0
    for key, nhom, ten in _ANTAXI_CP:
        v = get(key)
        if not v:
            continue
        tong_cp += v
        facts.append((cc, RT_CHIPHI, nhom, ten, v))
    if tong_cp:
        facts.append((cc, RT_HQKD, MA_CP, MA_CP, tong_cp))
    # PNLT: giữ ĐÚNG bộ dim1 bản THÁNG của An Taxi. Lưu ý An Taxi là đơn vị DUY NHẤT có dòng
    # giảm trừ tách riêng, và bản tháng ghi CẢ HAI tên cho DT GỘP ("Doanh thu bán hàng và cung
    # cấp dịch vụ" = tên đích danh cho bảng "Cấu trúc Doanh thu" ở revenue.py::_gross_of, và
    # "Doanh thu HH, DV" = tên chung cho KPI #1 ở overview.py) — cùng 1 giá trị, KHÔNG cộng đôi
    # vì 2 truy vấn lọc dim1_ilike khác nhau. Bỏ 1 trong 2 là mất số ở đúng 1 trong 2 chỗ đó.
    # ⚠ "Doanh thu HH, DV" của An Taxi là DT GỘP (mã 100), KHÁC XVP/HTX (= DT thuần).
    for key, rt, dim1 in (("lntt", RT_HQKD, MA_LNTT),
                          ("dt_gross", RT_PNLT, "Doanh thu bán hàng và cung cấp dịch vụ"),
                          ("dt_gross", RT_PNLT, "Doanh thu HH, DV"),
                          ("giam_tru", RT_PNLT, "Các khoản giảm trừ doanh thu"),
                          ("gia_von", RT_PNLT, "Giá vốn hàng bán"),
                          ("ln_gop", RT_PNLT, "Lợi nhuận gộp"),
                          ("dt_tai_chinh", RT_PNLT, "Doanh thu tài chính"),
                          ("tn_khac", RT_PNLT, "Thu nhập khác"),
                          ("lntt", RT_PNLT, "Lợi nhuận trước thuế"),
                          ("lnst", RT_PNLT, "Lợi nhuận sau thuế")):
        v = get(key)
        if v:
            facts.append((cc, rt, dim1, dim1, v))
    return facts


def _antaxi_facts(rows):
    """rows -> [(cost_center, report_type, dim1, dim3, value_VND)]. [] nếu sai layout.

    Họ file `.M.<YYYYMM>` — mỗi ngày MỘT SHEET. Họ còn lại xem `_bcqt_facts`."""
    cols, anchored, val = _kqkd_scan(rows, _CC_ANTAXI, _ANTAXI_ANCHOR)
    if not cols or "dt_thuan" not in anchored:
        return []
    facts = []
    for cc, j in cols:
        facts.extend(_antaxi_emit(cc, lambda k, _j=j: val(k, _j)))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "antaxi_bcqt" — An Taxi, sheet "BCQT PT." (mỗi NGÀY là một CỤM 3 CỘT trong CÙNG 1 sheet)
# ---------------------------------------------------------------------------------------------
# HỌ FILE THỨ HAI của CHÍNH thư mục `ANTAXI/baocaohqkdngay/`. Đọc TÊN file mà đoán họ là SAI —
# chữ `.M.`/`.D.` ngược hẳn nội dung:
#   · `B.7.AAG.TCKT.**M**.<YYYYMM>.xlsx`   "BC KQKD tạm"  — mỗi ngày 1 SHEET "1".."31"  -> "antaxi"
#   · `B.7.AAG.TCKT.**D**.<YYYYMMDD>.xlsx` "BC quản trị"  — mỗi ngày 1 CỤM CỘT          -> layout này
# Hai họ chung một thư mục nên chung một `_UNITS["ANTAXI"]` -> layout KHÔNG chọn được theo đơn vị.
# Chọn theo NỘI DUNG (có sheet "BCQT PT." không) qua `layout_phu`, cùng nguyên tắc `snapshot_sheet`
# của Trạm sạc và cùng lý do: hai họ nằm LẪN trong một thư mục, mốc theo ngày không tách được.
#
# ⚠ TÊN FILE LỆCH 1 NGÀY SO VỚI RUỘT. `.D.20260902.` chứa số đến hết 01/09; `.D.20260901.` RỖNG
# (đúng quy ước nguồn "16h45 ngày N+1 mới có báo cáo của ngày N"). TUYỆT ĐỐI không suy ngày số
# liệu từ tên file — ngày nằm ở nhãn cột dòng header ("01/09 - Sơn Tây").
#
# ⚠ MỖI FILE PHỦ TRỌN THÁNG (file ngày 11 có sẵn cột 01→10) nên khoá idempotent phải là KỲ, không
# phải file — y hệt lý do ở `period_source_key`. Và `cron_hqkdngay_daily.pick_targets` trả MỌI file
# khớp tháng rồi lặp theo (kỳ, công ty) chứ KHÔNG theo tên file, nên thứ tự nạp 12 file cùng đơn vị
# là tuỳ thứ tự metadata: nếu chỉ đọc file đang xét thì một file cũ chạy sau sẽ xoá sạch rồi ghi
# đè bằng bộ ngày NGẮN HƠN. Vì vậy mỗi lượt dựng lại CẢ KỲ từ TOÀN BỘ file — xem `_bcqt_per_day`.
# (`_BCQT_SHEET` khai sớm, cạnh `_UNITS`, vì được dùng làm khoá `layout_phu`.)

# Neo theo MÃ SỐ (cột "MÃ SỐ") — NGƯỢC với mọi layout khác trong file này (đều neo theo nhãn đã
# chuẩn hoá). Đảo ngược có lý do: ở sheet này mã số DUY NHẤT tuyệt đối (soát 57 dòng có mã, không
# mã nào trùng — khác hẳn HTX/XVP có 2 dòng cùng mã "10" nêu ở docstring đầu file), trong khi NHÃN
# có lỗi gõ ngay trên dòng LNTT ("LỢI NHUÂN TRƯỚC THUẾ TNDN"). Neo vào nhãn ở đây là neo vào một
# lỗi chính tả mà kế toán sửa lúc nào cũng được.
_BCQT_MA = {"dt_gross": "100", "giam_tru": "110", "dt_thuan": "120", "gia_von": "130",
            "ln_gop": "140", "cp_bien_doi": "150", "cp_co_dinh": "170", "dt_tai_chinh": "181",
            "cp_tai_chinh": "182", "tn_khac": "191", "cp_khac": "192", "lntt": "200",
            "lnst": "220"}

# Nhãn cột ngày: "01/09 - Sơn Tây" | "01/09 - Thái Nguyên" | "01/09 - Tổng".
_BCQT_COT = re.compile(r"^(\d{1,2})\s*/\s*(\d{1,2})\s*-\s*(.+)$")
# Dòng kỳ: "Từ ngày 01/09/2026 - Đến ngày 30/09/2026".
_BCQT_KY = re.compile(r"tu ngay\s*(\d{1,2})/(\d{1,2})/(\d{4}).*?den ngay\s*(\d{1,2})/(\d{1,2})/(\d{4})")


def _bcqt_ky(rows):
    """-> (tu_ngay, den_ngay) dạng 'YYYY-MM-DD', hoặc None nếu không thấy dòng kỳ.

    DÙNG ĐỂ LOẠI FILE SAI PHẠM VI, không phải để suy ngày số liệu. Ca thật (file
    `.D.20260912.` ngày 14/09/2026): cùng khuôn, cùng nhãn cột "01/09 - Sơn Tây", nhưng dòng kỳ
    ghi 01/01→31/12 và ruột là số THEO THÁNG (~3,1 tỷ/cột, TỔNG THÁNG = 25,5 tỷ = cả năm). Không
    có phép kiểm nào khác bắt được: mọi đẳng thức nội bộ của file đó vẫn khớp tuyệt đối.
    """
    for r in rows[:12]:
        for c in r or ():
            m = _BCQT_KY.search(_nd(c))
            if m:
                d1, m1, y1, d2, m2, y2 = (int(x) for x in m.groups())
                try:
                    return (f"{y1:04d}-{m1:02d}-{d1:02d}", f"{y2:04d}-{m2:02d}-{d2:02d}")
                except ValueError:
                    return None
    return None


def _bcqt_facts(rows, period):
    """rows của sheet "BCQT PT." -> [(ngay, [fact])]. [] nếu sai layout/sai kỳ.

    Chỉ trả những ngày CÓ SỐ: file dựng sẵn đủ 31 cụm cột cho cả tháng ngay từ đầu, các ngày chưa
    tới để trống/0. Header còn sinh cả "31/09" (tháng 9 không có ngày 31) nên ngày phải kiểm tính
    hợp lệ bằng `datetime.date`, không tin nhãn.
    """
    hdr_i = next((i for i, r in enumerate(rows[:14])
                  if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    y, mm = int(period[:4]), int(period[5:7])

    # (ngay, cost_center, j) — CỐ Ý BỎ cột "Tổng" của mỗi ngày: đã verify Sơn Tây + Thái Nguyên =
    # Tổng trên cả 10 ngày, ghi thêm dòng tổng chỉ tạo nguy cơ đếm đôi ở các truy vấn không đi qua
    # repository._per_file_resolved. Cùng lý do với `_kqkd_scan` bỏ cột "Tổng cộng".
    cot = []
    for j, c in enumerate(hdr):
        m = _BCQT_COT.match(str(c or "").strip())
        if not m:
            continue
        d, thang, ten = int(m.group(1)), int(m.group(2)), _nd(m.group(3))
        if thang != mm:
            continue
        try:
            ngay = datetime.date(y, thang, d).isoformat()
        except ValueError:                       # nhãn "31/09" do template sinh mù
            continue
        cc = next((cc for kw, cc in _CC_ANTAXI if ten == kw), None)
        if cc:
            cot.append((ngay, cc, j))
    if not cot:
        return []

    ma_j = next((j for j, c in enumerate(hdr) if _nd(c) == "ma so"), None)
    if ma_j is None:
        return []
    dong = {}
    for r in rows[hdr_i + 1:]:
        if not r or ma_j >= len(r) or r[ma_j] is None:
            continue
        dong[str(r[ma_j]).strip()] = r

    def val(key, j):
        r = dong.get(_BCQT_MA[key])
        return _num(r[j]) if r is not None and j < len(r) else None

    theo_ngay = {}
    for ngay, cc, j in cot:
        facts = _antaxi_emit(cc, lambda k, _j=j: val(k, _j))
        if facts:
            theo_ngay.setdefault(ngay, []).extend(facts)
    return sorted(theo_ngay.items())


def _bcqt_per_day(path, period, unit):
    """Dựng lại CẢ KỲ từ TOÀN BỘ file cùng kỳ trong thư mục -> (per_day, chẩn đoán).

    Gộp HAI NGUỒN nằm trong CÙNG một workbook nhưng có nhịp khác nhau:
      · P&L      — sheet `BCQT PT.`, mỗi ngày một cụm 3 cột, MỘT file phủ cả tháng.
      · Số dư    — `CDKT`/`CDPS`/`131`/`331`, mỗi file là ẢNH CHỤP MỘT NGÀY (tự khai "Từ ngày X
                   Đến ngày X"), nên phải đọc ĐỦ BỘ FILE mới thành chuỗi ngày.

    ⚠ HAI NGUỒN LỆCH NHAU MỘT NGÀY và đó KHÔNG phải lỗi: file `.D.20260915` có P&L đến hết 14/09
    nhưng số dư chốt 15/09. Vì vậy ngày cuối dải thường CHỈ CÓ số dư, chưa có P&L — không được
    lấy "thiếu P&L" làm cớ bỏ ngày đó, màn Công nợ/Tài sản sẽ mất đúng ngày mới nhất.

    ⚠ HAI PHẦN ĐƯỢC LOẠI RIÊNG. File `.D.20260912` có `BCQT PT.` hỏng (khai kỳ cả năm, ruột là số
    theo tháng) nhưng `CDKT`/`CDPS` của nó vẫn đúng ngày 12/09. Loại cả file là mất một mắt xích
    số dư và làm chuỗi đứt oan ở 12→14/09.

    File sau ĐÈ file trước theo từng ngày (không cộng) — các cột ngày của `BCQT PT.` bị ĐÓNG BĂNG
    ở lần xuất đầu nên mọi bản xuất khai y hệt nhau, đè là vô hại và khiến kết quả KHÔNG phụ
    thuộc thứ tự nạp (xem chú thích khối trên).
    """
    thu_muc = os.path.dirname(os.path.abspath(path))
    ten_sheet = unit.get("sheet_sodu", _SHEET_SODU_MAC_DINH)
    dung, bo_qua = [], []
    gop, gop_sodu = {}, {}
    for ten in sorted(os.listdir(thu_muc)):
        if not ten.lower().endswith(".xlsx") or ten.startswith("~$"):
            continue
        if _period_of(ten, True) != period:
            continue
        p = os.path.join(thu_muc, ten)
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        except Exception as e:
            bo_qua.append({"file": ten, "vi_sao": f"không mở được ({type(e).__name__})"})
            continue
        rows, sodu, ngay_sodu = None, [], None
        try:
            if _BCQT_SHEET not in (wb.sheetnames or []):
                continue                          # file họ `.M.` — không phải việc của layout này
            rows = [list(r) for r in wb[_BCQT_SHEET].iter_rows(values_only=True, max_row=80)]
            sn_cdkt = ten_sheet.get("cdkt")
            if sn_cdkt in (wb.sheetnames or []):
                ky_sd = _bcqt_ky([list(r) for r in
                                  wb[sn_cdkt].iter_rows(values_only=True, max_row=14)])
                # ĐÒI ĐÚNG MỘT NGÀY. `tu != den` nghĩa là bộ số dư đang ở khuôn luỹ kế/nhiều ngày
                # — lúc đó "số dư cuối kỳ" không còn là số dư của MỘT ngày, gán cho `den` là bịa.
                # Thà bỏ và báo còn hơn vẽ một chuỗi ngày sai mà không ai biết.
                if ky_sd and ky_sd[0] == ky_sd[1] and ky_sd[0][:7] == period:
                    ngay_sodu = ky_sd[0]
                    sodu = _snap_sodu_facts(wb, ten_sheet)
                elif ky_sd:
                    bo_qua.append({"file": ten, "phan": "số dư",
                                   "vi_sao": f"'{sn_cdkt}' khai {ky_sd[0]}..{ky_sd[1]}, không "
                                             f"phải ảnh chụp một ngày trong {period}"})
        finally:
            wb.close()
        if sodu and ngay_sodu:
            gop_sodu[ngay_sodu] = sodu
        ky = _bcqt_ky(rows)
        if ky and not (ky[0][:7] == period and ky[1][:7] == period):
            bo_qua.append({"file": ten, "phan": "P&L",
                           "vi_sao": f"'{_BCQT_SHEET}' khai kỳ {ky[0]}..{ky[1]}, không nằm trọn "
                                     f"trong {period}"})
            continue
        for ngay, facts in _bcqt_facts(rows, period):
            gop[ngay] = (facts, ten)
        dung.append(ten)

    per_day = [(n, list(gop.get(n, ((), None))[0]) + list(gop_sodu.get(n, ())))
               for n in sorted(set(gop) | set(gop_sodu))]
    diag = {"nguon_ngay": {"file_dung": dung, "so_file": len(dung),
                           "ngay_co_so_du": len(gop_sodu),
                           "ngay_co_pnl": len(gop)}}
    if bo_qua:
        diag["bo_qua_file"] = bo_qua
    # ĐỨT CHUỖI SỐ DƯ — số dư cuối ngày N phải bằng số dư đầu ngày N+1. Chỉ so HAI NGÀY LIỀN KỀ:
    # thiếu file ở giữa (kỳ 2026-09 không có file 13/09) thì lệch 12→14 là biến động THẬT của ngày
    # trống, báo đứt ở đó là báo oan. Chỉ soát loại có `du_dau` trong payload — BS_D/TS_D không
    # mang đầu kỳ nên `dk` luôn 0 và ngày nào cũng bị báo đứt.
    dut, truoc_ck, truoc_ngay = [], {}, None
    for ngay in sorted(gop_sodu):
        sodu = gop_sodu[ngay]
        lien_ke = truoc_ngay is not None and (datetime.date.fromisoformat(ngay)
                                              - datetime.date.fromisoformat(truoc_ngay)).days == 1
        for rt in {f[1] for f in sodu if f[6] and "du_dau" in f[6]}:
            ck = sum(f[4] for f in sodu if f[1] == rt)
            dk = sum((json.loads(f[6]).get("du_dau") or 0) * 1e9
                     for f in sodu if f[1] == rt and f[6])
            if lien_ke and truoc_ck.get(rt) is not None and abs(dk - truoc_ck[rt]) > 1:
                dut.append({"ngay": ngay, "loai": rt, "dau_ky": round(dk),
                            "cuoi_ky_hom_truoc": round(truoc_ck[rt]),
                            "lech": round(dk - truoc_ck[rt])})
            truoc_ck[rt] = ck
        truoc_ngay = ngay
    if dut:
        diag["so_du_dut_chuoi"] = dut
    # NGÀY MẤT CHI PHÍ — bất thường ĐÃ GẶP THẬT ở 01/09 và 08/09 kỳ 2026-09: có doanh thu nhưng
    # giá vốn chỉ còn dòng con "2.9 Giá vốn khác", ba nhóm chi phí lớn bằng 0 -> LNTT lộn DẤU
    # (+132 triệu thay vì −4,9 triệu). File tự nhất quán hoàn hảo nên không đẳng thức nào bắt
    # được; chỉ so tương quan doanh thu/chi phí mới thấy. VẪN NẠP (user chốt dùng nguồn này) —
    # cảnh báo để `cron_hqkdngay_daily` đưa lên bảng giám sát chứ không âm thầm.
    thieu = []
    for ngay, facts in per_day:
        # ĐÁNH CHỈ SỐ, KHÔNG GIẢI NÉN: từ khi gộp thêm cụm số dư thì `facts` lẫn HAI ĐỘ DÀI —
        # fact P&L 5 phần tử, fact số dư 7 (thêm dim2 + payload). `for _, rt, d1, _, v in facts`
        # ném ValueError ngay fact đầu tiên.
        dt = sum(f[4] for f in facts if f[1] == RT_HQKD and f[2] == MA_DT)
        cp = sum(f[4] for f in facts if f[1] == RT_HQKD and f[2] == MA_CP)
        if dt > 0 and cp < dt * 0.3:
            thieu.append({"ngay": ngay, "doanh_thu": round(dt), "tong_chi_phi": round(cp)})
    if thieu:
        diag["ngay_thieu_chi_phi"] = thieu
    return per_day, diag


# ---------------------------------------------------------------------------------------------
# Layout "tcode" — Global AI (GLOBALAI/baocaohqkdngay), mỗi ngày 1 sheet "D1".."D31"
# ---------------------------------------------------------------------------------------------
# Neo theo MÃ SỐ T-series (cột A, y hệt sheet THÁNG "HQKD HỢP NHẤT GA") — KHÁC layout "kqkd"/
# "antaxi" (neo theo NHÃN đã chuẩn hoá) vì GA không bị trùng mã như HTX/XVP, xem docstring đầu file.
# T201/T202/T203/T204 = CON TRỰC TIẾP của T200 (Σ 4 dòng = T200, verify sheet THÁNG khi có số thật)
# — cùng 4 nhóm CP mà bản THÁNG của GA hiện dùng qua sheet TT200 'KQKD' (mã 11/22/25/26 -> nhãn
# 'Giá vốn hàng bán'/'Chi phí tài chính'/'Chi phí bán hàng'/'Chi phí quản lý doanh nghiệp') để biểu
# đồ cơ cấu chi phí đọc được như nhau ở cả 2 chế độ Ngày/Tháng — xem agent_cli._TT200_CHITIEU.
# Chuẩn hoá nhãn Y HỆT `agent_cli._derive_kqkd_tseries._canon` (lý do xem trong `_tcode_facts`).
_TCODE_CANON = {"T201": "Giá vốn hàng bán", "T103": "Thu nhập khác"}


def _tcode_day_sheets(wb, period):
    """[(sheet_name, 'YYYY-MM-DD')] — sheet 'D1'..'D31' (file LUỸ KẾ: số sheet tăng dần theo ngày
    đã qua trong tháng, tên sheet KHÔNG zero-pad).

    ĐÃ THỬ RỒI VÀ ĐÃ BỎ (06/09/2026) — ĐỪNG LÀM LẠI: đường lùi "một file = một ngày", lấy ngày từ
    tên file `.D.20260905.` rồi đọc sheet 'BCHQKD'. Nghe hợp lý vì từ kỳ 2026-09 Trạm sạc/An Taxi
    gửi mỗi ngày một file và 'BCHQKD' vẫn đúng khuôn T-series `_tcode_facts` đọc được. NHƯNG SAI
    ĐƠN VỊ ĐO: bộ file đó KHÔNG phải báo cáo ngày, nó là MẪU BÁO CÁO THÁNG chụp lại theo ngày —
    so từng tên sheet thì file 'ngày' 05/09 giống TUYỆT ĐỐI file THÁNG `M.202607` và `M.202508`
    (11 sheet Trạm sạc / 9 sheet An Taxi), và số bên trong là LUỸ KẾ TỪ ĐẦU THÁNG:

        An Taxi, sheet KQKD, mã 01 (doanh thu):
          .D.20260901. "Từ 01/09 Đến 01/09"  148.773.050
          .D.20260902. "Từ 01/09 Đến 02/09"  317.258.324
          .D.20260903. "Từ 01/09 Đến 03/09"  431.011.417
          .D.20260904. "Từ 01/09 Đến 04/09"  532.069.771
          .D.20260905. "Từ 01/09 Đến 05/09"  632.844.149
        Coi mỗi file là một ngày rồi cộng 5 ngày -> 2.061.956.711 đ, PHÓNG ĐẠI 3,26 LẦN so với
        số thật 632.844.149 đ của cả kỳ.

    Mẫu ngày THẬT thì ngược lại — không luỹ kế: file `.D.20260801.` của Trạm sạc có T101 ở D4 =
    343.874.132, D6 = 933.268.519, D27 = 143.829.479 (D27 < D6), Σ 31 sheet = đúng số tháng.

    Vì vậy hàm này CHỈ nhận dải sheet ngày. Bộ file mẫu-tháng-chụp-theo-ngày phải đi đường khác:
    số THÁNG thì lấy snapshot MỚI NHẤT của kỳ (nó chính là luỹ kế đến ngày đó), số NGÀY thì lấy từ
    sheet sổ nhật ký trong chính file ('DATA' của An Taxi · 'Sổ nhật ký chung' của Trạm sạc) — sheet
    đó chứa bút toán CỦA ĐÚNG NGÀY ĐÓ, có cả vế chi phí mà sheet P&L đang để trống."""
    y, mm = int(period[:4]), int(period[5:7])
    out = []
    for s in wb.sheetnames:
        m = re.fullmatch(r"[Dd](\d{1,2})", s.strip())
        if m and 1 <= int(m.group(1)) <= 31:
            out.append((s, f"{y:04d}-{mm:02d}-{int(m.group(1)):02d}"))
    return sorted(out, key=lambda x: x[1])


def _tcode_byco(rows):
    """rows -> {mã T-series: (nhãn, giá trị)}. {} nếu sai layout (không dò ra header 'Mã số').

    TÁCH RA KHỎI `_tcode_facts` (11/09/2026) để luồng SNAPSHOT LUỸ KẾ dùng lại được: nó cần đọc
    giá trị THÔ của nhiều file rồi mới trừ nhau ra số ngày, chứ không dựng fact ngay từ 1 sheet.
    Logic giữ NGUYÊN VẸN.

    Sheet 'BCHQKD' của bản tự động ghi header cột giá trị là 'T9' (số THÁNG) nên không khớp mẫu
    'd\\d{2}' — nó rơi về `ten_j + 1`, và đó ĐÚNG là cột D. Đã đo: nới mẫu để bắt cả 'T9' cho ra
    kết quả GIỐNG HỆT trên cả 10 mã của file 10/09, nên không nới."""
    hdr_i = next((i for i, r in enumerate(rows[:10]) if any(_nd(c) == "ma so" for c in r if c is not None)), None)
    if hdr_i is None:
        return {}
    hdr = rows[hdr_i]
    ma_j = next(j for j, c in enumerate(hdr) if _nd(c) == "ma so")
    ten_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("chi tieu")), ma_j + 1)
    val_j = next((j for j, c in enumerate(hdr) if j not in (ma_j, ten_j)
                  and re.fullmatch(r"d\d{2}", _nd(c))), ten_j + 1)

    byco = {}
    for r in rows[hdr_i + 1:]:
        c = str(r[ma_j]).strip() if ma_j < len(r) and r[ma_j] not in (None, "") else ""
        if re.fullmatch(r"T\d{3}", c) and c not in byco:
            lab = str(r[ten_j]).strip() if ten_j < len(r) and r[ten_j] not in (None, "") else c
            byco[c] = (lab, _num(r[val_j]) if val_j < len(r) else None)
    return byco


def _tcode_facts(rows, profit_pnlt=("Lợi nhuận trước thuế",), pnlt_skip=()):
    """rows -> [(None, report_type, dim1, dim3, value_VND)] cho 1 ngày. [] nếu sai layout.
    Cột giá trị dò theo header khớp 'D\\d{2}' (zero-pad) — TỰ suy ra cột đúng của sheet này, không
    cần biết trước số ngày, vì mỗi sheet chỉ có DUY NHẤT 1 cột như vậy.

    `profit_pnlt`: (các) tên dim1 PNLT ghi từ T300 — THAM SỐ HOÁ vì 2 đơn vị dùng chung mã T-series
    nhưng quy ước LNTT/LNST bản THÁNG KHÁC NHAU (xem docstring khối 'tcode' + đơn vị TRAMSAC):
    GA ghi 'Lợi nhuận trước thuế' (chưa chắc = LNST thật khi phát sinh thuế); Trạm sạc bản THÁNG
    CHỈ có đúng 1 dim1 PNLT 'Lợi nhuận sau thuế' (không có 'Lợi nhuận trước thuế' riêng, verify DB
    2026-06: PNLT 'Lợi nhuận sau thuế' = HQKD 1112 = -0,3844, bằng nhau tuyệt đối -> T300 ở Trạm
    sạc ĐÃ LÀ LNST, gán nhãn 'Lợi nhuận trước thuế' cho nó sẽ SAI tên cột trên bảng Cấu trúc CP."""
    return _tcode_facts_byco(_tcode_byco(rows), profit_pnlt, pnlt_skip)


def _tcode_facts_byco(byco, profit_pnlt=("Lợi nhuận trước thuế",), pnlt_skip=()):
    """{mã: (nhãn, giá trị)} -> [(None, report_type, dim1, dim3, value_VND)]. Xem `_tcode_facts`."""
    if "T101" not in byco or "T200" not in byco or "T300" not in byco:
        return []

    def val(code):
        return byco.get(code, (None, None))[1]

    facts = []
    dt = val("T101")
    if dt:
        facts.append((None, RT_HQKD, MA_DT, MA_DT, dt))
        facts.append((None, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
        facts.append((None, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
    if val("T200"):
        facts.append((None, RT_HQKD, MA_CP, MA_CP, val("T200")))
    lntt = val("T300")
    if lntt:
        facts.append((None, RT_HQKD, MA_LNTT, MA_LNTT, lntt))
        for nm in profit_pnlt:
            facts.append((None, RT_PNLT, nm, nm, lntt))
    # Lợi nhuận gộp = T100 − T201 — BỔ SUNG 2026-08-06. Bản THÁNG của CẢ HAI đơn vị đều có dim1 này
    # (verify DB: Trạm sạc T07 'Lợi nhuận gộp' = 1,297842 · GA T06 có 'Lợi nhuận gộp'), bản ngày
    # trước đây thiếu -> thẻ LN gộp trống ở chế độ Ngày. LƯU Ý: quyết định cũ "GA không suy diễn
    # LNST/LN gộp" chỉ đúng cho LNST (T300 chưa chắc đã trừ thuế TNDN — GA bản tháng có dòng 'Chi
    # phí thuế TNDN' riêng); LN gộp = T100 − T201 KHÔNG liên quan thuế nên áp cho cả hai.
    #
    # SỬA 11/08/2026 — lấy T101 (doanh thu thuần) chứ KHÔNG phải T100. `T100 = T101 + T102 (DT tài
    # chính) + T103 (thu nhập khác)`; chính hàm này đã dùng `dt = val("T101")` cho doanh thu ở trên,
    # riêng dòng lãi gộp lấy T100 -> trộn doanh thu tài chính vào LÃI GỘP. Bản THÁNG
    # (`agent_cli._derive_kqkd_tseries`) sửa cùng lượt để hai bản không lệch định nghĩa.
    if val("T201") is not None and dt is not None and (dt - val("T201")):
        facts.append((None, RT_PNLT, "Lợi nhuận gộp", "Lợi nhuận gộp", dt - val("T201")))
    # MỌI mã còn lại -> PNLT giữ NHÃN GỐC, chuẩn hoá 2 nhãn y hệt bản THÁNG
    # (`agent_cli._derive_kqkd_tseries._canon`): T201 -> 'Giá vốn hàng bán' (nguồn gõ 'Gía vốn' làm
    # metrics lọc ILIKE '%giá vốn%' ACCENT-SENSITIVE trượt) · T103 -> 'Thu nhập khác'.
    # `pnlt_skip`: GA bản tháng chạy extractor TT200 (KHÔNG phải T-series) nên KHÔNG có dim1
    # 'Doanh thu bán hàng' (T101 nhãn gốc) — thêm vào sẽ đẻ chỉ tiêu bản tháng không có, trùng giá
    # trị với 'Doanh thu HH, DV'. Trạm sạc bản tháng CÓ dòng đó -> không skip.
    for code, (lab, v) in byco.items():
        if code in ("T100", "T200", "T300") or code in pnlt_skip or not v:
            continue
        ten = _TCODE_CANON.get(code, lab)
        facts.append((None, RT_PNLT, ten, ten, v))
    # CHIPHI: mã con TRỰC TIẾP của T200 = '^T2\d{2}$' trừ T200, NHÃN GỐC (bản tháng dùng nhãn gốc —
    # verify DB Trạm sạc T07: CHIPHI 'Gía vốn hàng bán' GIỮ typo, trong khi PNLT là bản chuẩn hoá;
    # bản ngày trước đây hardcode nhãn chuẩn -> cùng 1 khoản bị tách 2 nhóm khác tên giữa Tháng/Ngày).
    # Chỉ dựng khi Σ con PHỦ HẾT T200 (1%) — y hệt điều kiện `_covers` bản tháng, tránh breakdown thiếu.
    cp = [(byco[c][0], byco[c][1]) for c in byco
          if re.fullmatch(r"T2\d{2}", c) and c != "T200" and byco[c][1]]
    if cp and val("T200") and abs(sum(x[1] for x in cp) - val("T200")) <= abs(val("T200")) * 0.01:
        for nhom, v in cp:
            facts.append((None, RT_CHIPHI, nhom, nhom, v))
    return facts


# ---------------------------------------------------------------------------------------------
# Luồng "snapshot luỹ kế" — MỘT FILE = MỘT ẢNH CHỤP LUỸ KẾ TỪ ĐẦU THÁNG (Trạm sạc, từ kỳ 2026-09)
# ---------------------------------------------------------------------------------------------
# ĐÂY KHÔNG PHẢI đường lùi đã bị bác ở docstring `_tcode_day_sheets`. Cái bị bác là "coi MỖI FILE
# LÀ MỘT NGÀY rồi cộng lại" — cộng 5 snapshot luỹ kế của An Taxi ra 2.061.956.711 đ trong khi số
# thật cả kỳ là 632.844.149 đ (phóng đại 3,26 lần). Ở đây làm NGƯỢC LẠI: đọc cả bộ snapshot của
# kỳ rồi TRỪ LIÊN TIẾP ra số phát sinh riêng từng ngày, nên Σ mọi ngày = snapshot CUỐI CÙNG, đúng
# bằng số luỹ kế của kỳ. Đo trên bộ file thật 01→09/09/2026 (mapping `Mapping_TramSac_daily.xlsx`,
# 3 cột "Đường dẫn BC tự động" / "Sheet cần lấy" / "Mô tả dữ liệu lấy trên BC tự động"):
#     ô D7 (T100) qua các ngày: 0 · 0 · 0 · 0 · 0 · 0 · 29.936.027 · 282.317.354 · 282.317.354
#     -> số NGÀY: 07/09 = 29.936.027 · 08/09 = 252.381.327 · 09/09 = 0 · Σ = 282.317.354 ✓
# Bằng chứng cột D là luỹ kế chứ không phải số ngày: ô D6 ghi "T9" (số THÁNG), và BCKQKD/BCĐKT/
# BC LCTT trong CÙNG file đều ghi header "Từ ngày 01/09/2026 - Đến ngày 09/09/2026".
#
# CHỈ BẬT CHO ĐƠN VỊ KHAI `snapshot_sheet` trong `_UNITS`, và CHỈ khi file KHÔNG có dải sheet ngày
# "D1".."D31". GA cũng layout "tcode" và tên file cũng có 8 số (".D.20260601.") nhưng ruột vẫn là
# 31 sheet ngày -> không khai `snapshot_sheet`, không đi nhánh này.
_SNAP_NGAY_SHEETS = ("BCKQKD", "BCĐKT", "BC LCTT")


def _snap_day_of(file_name):
    """'B.3.TC.TCKT.D.20260909.Baocaotaichinhrieng.xlsx' -> '2026-09-09'. None nếu tên thiếu ngày.

    Tên file Trạm sạc lúc có lúc không có dấu cách sau 'B.3.TC.' — `_period_of` đã chịu được, ở
    đây chỉ cần 8 chữ số sau '.D.' nên cũng không vướng."""
    m = re.search(r"\.D\.(\d{4})(\d{2})(\d{2})\D", file_name)
    if not m:
        return None
    y, mm, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f"{y:04d}-{mm:02d}-{dd:02d}" if 1 <= mm <= 12 and 1 <= dd <= 31 else None


def _snap_ky(wb):
    """(tu_ngay, den_ngay) dạng 'YYYY-MM-DD', đọc dòng "Từ ngày dd/mm/yyyy - Đến ngày dd/mm/yyyy"
    ở đầu BCKQKD/BCĐKT/BC LCTT. (None, None) nếu không đọc được.

    ĐÂY LÀ CHỖ FILE TỰ KHAI NÓ LÀ LOẠI GÌ, và là khoá của cả luồng snapshot:
        tu == den            -> báo cáo RIÊNG NGÀY   -> dùng thẳng số trong file
        tu == mùng 1 < den   -> báo cáo LUỸ KẾ       -> phải trừ snapshot trước ra số ngày

    KHÔNG ĐƯỢC SUY TỪ TÊN FILE. Ngày trong tên chỉ là ngày kéo, và bản chất kỳ đã đổi NGAY GIỮA
    CHỪNG mà tên file không đổi một ký tự: tới 10/09/2026 bộ file còn ghi "Từ 01/09 - Đến 09/09"
    (luỹ kế), rạng sáng 11/09 kế toán xuất lại toàn bộ thành "Từ 08/09 - Đến 08/09" (riêng ngày).
    Bám tên file thì lần đổi đó đi qua lặng lẽ và mọi con số thành rác — đã xảy ra thật một lượt.
    Đọc không ra kỳ thì BỎ QUA file và nói ra, tuyệt đối không đoán: đoán sai ở đây không lộ ra
    thành lỗi, nó lộ ra thành số sai trên dashboard."""
    for sn in _SNAP_NGAY_SHEETS:
        if sn not in wb.sheetnames:
            continue
        for r in wb[sn].iter_rows(min_row=1, max_row=12, values_only=True):
            for c in r:
                if not isinstance(c, str):
                    continue
                m = re.search(r"[Tt]ừ\s*ngày\s*:?\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*[-–]?\s*"
                              r"[Đđ]ến\s*ngày\s*:?\s*(\d{1,2})/(\d{1,2})/(\d{4})", c)
                if m:
                    g = [int(x) for x in m.groups()]
                    return (f"{g[2]:04d}-{g[1]:02d}-{g[0]:02d}",
                            f"{g[5]:04d}-{g[4]:02d}-{g[3]:02d}")
    return None, None


# ---------------------------------------------------------------------------------------------
# SỐ DƯ THEO NGÀY — công nợ phải thu / phải trả từ "Sổ tổng hợp CN ..." của chính file ngày
# ---------------------------------------------------------------------------------------------
# KHÁC HẲN P&L Ở MỘT ĐIỂM QUYẾT ĐỊNH: đây là SỐ DƯ (stock), không phải phát sinh (flow). Doanh thu
# 12 ngày cộng lại ra doanh thu tháng; số dư 12 ngày cộng lại thì VÔ NGHĨA (gấp ~12 lần). Vì vậy:
#   · deriver ghi số dư CUỐI NGÀY của từng ngày, KHÔNG trừ liên tiếp như nhánh luỹ kế;
#   · phía đọc (app/metrics) phải lấy DÒNG CỦA NGÀY CUỐI trong khoảng, KHÔNG được SUM.
# Hai vế phải đi cùng nhau — nạp mà quên vế đọc là công nợ phình lên gấp số ngày.
#
# Quy ước dim1/amount/payload GIỮ ĐÚNG bản THÁNG (`agent_cli._derive_congno`, verify raw_rows
# Trạm sạc 2026-07): dim1 = tên đối tượng · amount = số dư RÒNG cuối kỳ (tỷ) · payload mang
# du_dau/ps_tang/ps_giam/ma_dt. Lệch quy ước là bảng Top 10 công nợ hiện hai nhóm tên khác nhau
# giữa chế độ Tháng và Ngày.
#   TK131 phải thu (dư NỢ): dư = Nợ − Có · tăng = PS Nợ · giảm = PS Có
#   TK331 phải trả (dư CÓ): dư = Có − Nợ · tăng = PS Có · giảm = PS Nợ
_CONGNO_SHEET = {
    "PTHU": ("Sổ tổng hợp CN phải thu", "no"),
    "PTRA": ("Sổ tổng hợp CN phải trả", "co"),
}

# TÊN SHEET CỦA CỤM SỐ DƯ — khác nhau giữa các đơn vị dù RUỘT GIỐNG HỆT (cùng mẫu S31-DN cho sổ
# công nợ, S06-DN cho bảng cân đối phát sinh, B01-DN cho CĐKT). Vì thế tên nằm ở cấu hình đơn vị
# (`_UNITS[...]["sheet_sodu"]`), còn các hàm bóc thì dùng chung không sửa một dòng:
#   · Trạm sạc : "Sổ tổng hợp CN phải thu" · "Sổ tổng hợp CN phải trả" · "BCĐKT" · "BCĐPS"
#   · An Taxi  : "131"                     · "331"                     · "CDKT"  · "CDPS"
_SHEET_SODU_MAC_DINH = {"pthu": _CONGNO_SHEET["PTHU"][0], "ptra": _CONGNO_SHEET["PTRA"][0],
                        "cdkt": "BCĐKT", "cdps": "BCĐPS"}


# TẦNG DƯỚI GHI THEO NGHIỆP VỤ THAY VÌ NỢ/CÓ (khối Dự án, 20/09/2026). Bảng tổng hợp công nợ của
# mẫu BCTC mà kế toán Dự án dùng không ghi "Nợ"/"Có" mà ghi thẳng việc: sheet `131` để
# "PHẢI THU | ĐÃ THU", sheet `331` để "ĐÃ TRẢ | PHẢI TRẢ" (chú ý THỨ TỰ CỘT NGƯỢC NHAU giữa hai
# sheet — nên phải ánh xạ theo NHÃN, không theo vị trí).
#
# Ánh xạ dưới đây là định nghĩa kế toán, không phụ thuộc chiều của report_type: phát sinh phải thu
# ghi bên NỢ của TK 131, thu tiền ghi bên CÓ; phải trả ghi bên CÓ của TK 331, trả tiền ghi bên NỢ.
# Đối chứng trên file 18/09: sheet 131 dòng CỘNG có cuối kỳ Nợ 37.378.363.809 − Có 24.660.179.191
# = 12.718.184.618, khớp đúng ô "Phải thu KH ngắn hạn" của CĐKT.
#
# KHÔNG đụng 5 đơn vị đang chạy: chúng ghi "Nợ"/"Có" nên rơi vào hai nhánh `startswith` phía trên,
# bảng này chỉ được hỏi tới khi cả hai nhánh đó trượt.
_BEN_THEO_NGHIEP_VU = {"phai thu": "no", "da thu": "co", "phai tra": "co", "da tra": "no"}


def _snap_2tang(rows):
    """Dò header 2 tầng Nợ/Có -> (dòng header trên, {dk_no/dk_co/ps_no/ps_co/ck_no/ck_co: cột}).

    Dùng chung cho "Sổ tổng hợp CN ..." (mẫu S31-DN) và "BCĐPS" (mẫu S06-DN): hai biểu mẫu khác
    nhau nhưng cùng bố cục — tầng trên gom cụm Đầu kỳ / Phát sinh / Cuối kỳ (chỉ ô ĐẦU mỗi cụm có
    nhãn, ô sau bị merge nên phải loang sang phải), tầng dưới ghi Nợ / Có. Dò theo NỘI DUNG chứ
    không gán cứng cột C..H vì số cột đệm đổi theo kỳ; so bằng `in` chứ không `==` vì nhãn hai
    sheet khác nhau ("Đầu kỳ" vs "Số dư đầu kỳ"). Trả (None, {}) nếu không dựng đủ 6 cột."""
    def cum_of(c):
        k = _nd(c)
        if not k:
            return None
        return "dk" if "dau ky" in k else "ps" if "phat sinh" in k else "ck" if "cuoi ky" in k else None

    top = next((i for i, r in enumerate(rows[:20]) if sum(1 for c in r if cum_of(c)) >= 2), None)
    if top is None or top + 1 >= len(rows):
        return None, {}
    cum, nhan = {}, None
    for j, c in enumerate(rows[top]):
        nhan = cum_of(c) or nhan
        if nhan:
            cum[j] = nhan
    col = {}
    for j, c in enumerate(rows[top + 1]):
        v = _nd(c)
        ben = ("no" if v.startswith("no") else "co" if v.startswith("co")
               else _BEN_THEO_NGHIEP_VU.get(v))
        if j in cum and ben:
            col[f"{cum[j]}_{ben}"] = j
    need = {"dk_no", "dk_co", "ps_no", "ps_co", "ck_no", "ck_co"}
    return (top, col) if need <= set(col) else (None, {})


# TSNV = TOÀN BỘ dòng của BCĐKT (bảng Tài sản / Nguồn vốn của màn `finance`, chỉ tiêu 27–28 của
# mapping). KHÁC `BS_D` — BS chỉ có 3 mã tổng cho thẻ KPI, còn đây là cả bảng để vẽ cơ cấu.
# `dim1` chỉ gán cho dòng CẤP LA MÃ + mã 400, y hệt bản THÁNG (verify raw_rows 2026-07: đúng
# 5+6+2+1 = 14 dòng có nhóm, 103 dòng còn lại để rỗng). Gán rộng hơn là bảng cơ cấu cộng đôi vì
# dòng cha và dòng con cùng lọt một nhóm.
_TSNV_NHOM = {**{m: "TS ngắn hạn" for m in ("110", "120", "130", "140", "150")},
              **{m: "TS dài hạn" for m in ("210", "220", "230", "240", "250", "260")},
              **{m: "Nợ phải trả" for m in ("310", "330")}, "400": "Vốn chủ"}


def _snap_tsnv_facts(rows, mau=None):
    """rows sheet BCĐKT -> [fact] TSNV_D: mỗi dòng CĐKT một fact. `mau` xem `_MAU_CDKT`."""
    mau = mau or _MAU_CDKT["b01dn"]
    hdr = next((i for i, r in enumerate(rows[:20])
                if any(_nd(c) == "ma so" for c in r)
                and any(any(k in _nd(c) for k in mau["nhan"]) for c in r)), None)
    if hdr is None:
        return []
    ma_j = next(j for j, c in enumerate(rows[hdr]) if _nd(c) == "ma so")
    ten_j = next((j for j, c in enumerate(rows[hdr]) if _nd(c) in ("tai san", "chi tieu")), 0)
    val_j = next((j for j, c in enumerate(rows[hdr])
                  if any(k in _nd(c) for k in mau["nhan"])), None)
    if val_j is None:
        return []
    facts, da_co = [], set()
    for r in rows[hdr + 1:]:
        ma = str(r[ma_j]).strip() if ma_j < len(r) and r[ma_j] not in (None, "") else ""
        ten = str(r[ten_j]).strip() if ten_j < len(r) and r[ten_j] not in (None, "") else ""
        if not re.fullmatch(r"\d{3}", ma) or not ten or ma in da_co:
            continue
        da_co.add(ma)
        v = _num(r[val_j]) if val_j < len(r) else None
        pl = json.dumps({"ps_tang": 0.0, "ps_giam": 0.0, "ma_so": ma,
                         "unit": "ty", "grain": "day"}, ensure_ascii=False)
        facts.append((None, "TSNV_D", mau["nhom"].get(ma, ""), "", v or 0.0, ten, pl))
    return facts


def _snap_congno_facts(rows, rt, chieu):
    """rows sheet sổ tổng hợp công nợ -> [(None, rt, dim1, dim3, số dư tỷ, dim2, payload)]."""
    top, col = _snap_2tang(rows)
    if top is None:
        return []
    # Nhãn cột MÃ / TÊN đối tượng khác nhau giữa hai mẫu: S31-DN ghi "Mã ĐT"/"Tên ĐT", còn mẫu của
    # kế toán Dự án ghi "MÃ"/"TÊN CÔNG NỢ PHẢI THU" (hoặc "...PHẢI TRẢ"). Dò cả hai bộ nhãn, đừng
    # để rơi xuống mặc định: cột 0 của mẫu Dự án là ô trống và cột kế là số thứ tự, nên `dim1` sẽ
    # thành "1","2","3"… — một danh sách "khách hàng" toàn số mà không có lỗi nào được ném ra.
    ma_j = next((j for j, c in enumerate(rows[top])
                 if _nd(c).startswith("ma dt") or _nd(c) == "ma"), 0)
    ten_j = next((j for j, c in enumerate(rows[top])
                  if _nd(c).startswith(("ten dt", "ten cong no"))), ma_j + 1)

    pos, neg = ("no", "co") if chieu == "no" else ("co", "no")
    facts = []
    for r in rows[top + 2:]:
        ten = str(r[ten_j]).strip() if ten_j < len(r) and r[ten_j] not in (None, "") else ""
        if not ten:
            continue
        g = lambda k: (r[col[k]] if col[k] < len(r) and isinstance(r[col[k]], (int, float)) else 0)  # noqa: E731
        dau = g(f"dk_{pos}") - g(f"dk_{neg}")
        tang, giam = g(f"ps_{pos}"), g(f"ps_{neg}")
        cuoi = g(f"ck_{pos}") - g(f"ck_{neg}")
        if not (dau or tang or giam or cuoi):
            continue
        ma = str(r[ma_j]).strip() if ma_j < len(r) and r[ma_j] not in (None, "") else ""
        pl = json.dumps({"du_dau": round(dau * 1e-9, 9), "ps_tang": round(tang * 1e-9, 9),
                         "ps_giam": round(giam * 1e-9, 9), "ma_dt": ma,
                         "unit": "ty", "grain": "day"}, ensure_ascii=False)
        facts.append((None, rt, ten, ten, cuoi, None, pl))
    return facts


# BCĐKT -> BS (3 mã tổng) + TS (TSCĐ). Dò theo MÃ ở cột "Mã số", KHÔNG theo số dòng: mapping
# `Mapping_TramSac_daily.xlsx` ghi mã 300 ở dòng 73 trong khi file thật để ở dòng 74 (dòng 73 là
# mã 270) — gán cứng dòng là lấy nhầm tổng tài sản làm nợ phải trả.
_CDKT_BS = ("270", "300", "400")          # tổng tài sản · nợ phải trả · vốn chủ sở hữu
_CDKT_TS = {"gtcl": "221", "nguyen_gia": "222", "hao_mon": "223"}

# MẪU BIỂU CĐKT — hai mẫu, KHÁC NHAU CẢ NHÃN CỘT LẪN BẢNG MÃ (19/09/2026).
#
# Trước đó mọi thứ gán cứng theo B01-DN (TT200). Hai HTX Xanh nộp mẫu **B01-HTX** (TT 71/2024,
# "BÁO CÁO TÌNH HÌNH TÀI CHÍNH") nên không đọc được gì: cột giá trị ghi "Kỳ này" chứ không "Số
# cuối kỳ", và mã số mang nghĩa khác — tổng tài sản là **200** chứ không phải 270, TSCĐ là
# **150/151/152** chứ không phải 221/222/223. Dò bằng bảng của B01-DN thì hoặc trả rỗng (sai nhãn),
# hoặc tệ hơn là khớp nhầm mã và gán nhóm sai cho cả bảng cơ cấu.
#
# `bs` map mã NGUỒN -> mã CHUẨN: dòng `BS_D` bắt buộc mang dim1 "270"/"300"/"400" vì phía đọc hỏi
# đúng ba mã đó (`finance.py`: `snapshot_sum(ds,"BS",to,dim1_in=["270"])`). Đổi mã ở đây là thẻ
# Tổng tài sản trắng mà không có lỗi nào nổ.
#
# `nhom` CHỈ gán cho dòng CẤP TỔNG, không gán cho dòng con — cùng lý do như bản B01-DN: cha và con
# cùng lọt một nhóm là bảng cơ cấu cộng đôi. Với B01-HTX: 150 là cha của 151/152, 300 là cha của
# 310..380, 400 là cha của 410..440.
#
# ⚠ B01-HTX KHÔNG TÁCH NGẮN/DÀI HẠN — mẫu rút gọn chỉ liệt kê 8 khoản tài sản liền mạch. Việc xếp
# 150 (TSCĐ) và 160 (TS chung không chia) vào "TS dài hạn", phần còn lại vào "TS ngắn hạn" là một
# DIỄN GIẢI của mình theo bản chất khoản mục, không phải thứ nguồn khai. Hai nhóm này chỉ dùng để
# vẽ thanh cơ cấu; tổng thì lấy thẳng mã 200 nên không phụ thuộc cách xếp.
_MAU_CDKT = {
    "b01dn": {
        "nhan": ("cuoi ky",),
        "bs": {"270": "270", "300": "300", "400": "400"},
        "ts": _CDKT_TS,
        "nhom": {**{m: "TS ngắn hạn" for m in ("110", "120", "130", "140", "150")},
                 **{m: "TS dài hạn" for m in ("210", "220", "230", "240", "250", "260")},
                 **{m: "Nợ phải trả" for m in ("310", "330")}, "400": "Vốn chủ"},
    },
    "b01htx": {
        # "Kì này" (i ngắn) là lỗi gõ CÓ THẬT ở bản XTQ 15/09, bản 16-17 lại ghi "Kỳ này" — nhận
        # cả hai, không thì mất đúng một ngày vì một dấu.
        "nhan": ("ky nay", "ki nay"),
        "bs": {"200": "270", "300": "300", "400": "400"},
        "ts": {"gtcl": "150", "nguyen_gia": "151", "hao_mon": "152"},
        "nhom": {**{m: "TS ngắn hạn" for m in ("110", "120", "130", "140", "170", "180")},
                 **{m: "TS dài hạn" for m in ("150", "160")},
                 "300": "Nợ phải trả", "400": "Vốn chủ"},
    },
}


def _cdkt_doc(rows, mau):
    """(dòng header, {mã: giá trị}) của sheet CĐKT theo mẫu. (None, {}) nếu không dò ra cột giá trị."""
    def co_nhan(r):
        return any(any(k in _nd(c) for k in mau["nhan"]) for c in r)

    hdr = next((i for i, r in enumerate(rows[:20])
                if any(_nd(c) == "ma so" for c in r) and co_nhan(r)), None)
    if hdr is None:
        return None, {}
    ma_j = next(j for j, c in enumerate(rows[hdr]) if _nd(c) == "ma so")
    val_j = next((j for j, c in enumerate(rows[hdr])
                  if any(k in _nd(c) for k in mau["nhan"])), None)
    if val_j is None:
        return None, {}
    byma = {}
    for r in rows[hdr + 1:]:
        m = str(r[ma_j]).strip() if ma_j < len(r) and r[ma_j] not in (None, "") else ""
        if re.fullmatch(r"\d{3}", m) and m not in byma:
            byma[m] = _num(r[val_j]) if val_j < len(r) else None
    return hdr, byma


def _snap_cdkt_facts(rows, mau=None):
    """rows sheet BCĐKT -> [fact] cho BS_D và TS_D (số dư CUỐI KỲ). `mau` xem `_MAU_CDKT`."""
    mau = mau or _MAU_CDKT["b01dn"]
    hdr, byma = _cdkt_doc(rows, mau)
    if hdr is None:
        return []
    facts = [(None, "BS_D", chuan, chuan, byma[nguon], None, None)
             for nguon, chuan in mau["bs"].items() if byma.get(nguon) is not None]
    ts = mau["ts"]
    g = byma.get(ts["gtcl"])
    if g is not None:
        # hao_mon lưu DƯƠNG (bản THÁNG verify 2026-07: hao_mon 0,5791 trong khi mã 223 ghi âm).
        # GTCL lấy thẳng mã gtcl, KHÔNG tính nguyên giá − hao mòn: cột hao mòn đã âm sẵn nên phép
        # trừ ra sai gấp đôi.
        pl = json.dumps({"nguyen_gia": round((byma.get(ts["nguyen_gia"]) or 0) * 1e-9, 9),
                         "hao_mon": round(abs(byma.get(ts["hao_mon"]) or 0) * 1e-9, 9),
                         "khau_hao_ky": 0.0, "het_kh_con_sd": 0.0, "het_kh_thanh_ly": 0.0,
                         "unit": "ty", "grain": "day"}, ensure_ascii=False)
        facts.append((None, "TS_D", "TSCĐ (theo CĐKT)", "TSCĐ (theo CĐKT)", g, None, pl))
    return facts


# BCĐPS -> THUE (TK 133 / 333x) + HH (TK 15x). Nhãn giữ ĐÚNG bản THÁNG (verify raw_rows 2026-07).
_THUE_TK = {"133": ("Phải thu", "Thuế GTGT được khấu trừ", "no"),
            "3331": ("Phải nộp", "Thuế Giá trị gia tăng (GTGT)", "co"),
            "3334": ("Phải nộp", "Thuế Thu nhập doanh nghiệp (TNDN)", "co"),
            "3335": ("Phải nộp", "Thuế Thu nhập cá nhân (TNCN)", "co"),
            "3338": ("Phải nộp", "Thuế, phí khác", "co")}
_HH_TK = {"151": "Hàng mua đang đi đường", "152": "Nguyên liệu, vật liệu",
          "153": "Công cụ, dụng cụ", "154": "Chi phí sản xuất kinh doanh dở dang",
          "155": "Thành phẩm", "156": "Hàng hóa"}


# TK công nợ ở mức TỔNG, CHỈ dùng cho đơn vị KHÔNG có sổ chi tiết theo đối tượng (2 HTX Xanh,
# 19/09/2026). Đơn vị CÓ sổ (Trạm sạc, An Taxi, An KS) mà bật cái này là cộng đôi: một lần theo
# từng đối tượng, một lần nữa ở dòng tổng. Vì vậy `_snap_sodu_facts` chỉ truyền `congno_tong=True`
# khi `sheet_sodu` của đơn vị KHÔNG khai `pthu`/`ptra`.
_CONGNO_TK = {"131": ("PTHU_D", "no", "Phải thu khách hàng (tổng)"),
              "331": ("PTRA_D", "co", "Phải trả nhà cung cấp (tổng)")}


def _snap_cdps_facts(rows, congno_tong=False):
    """rows sheet BCĐPS -> [fact] cho THUE_D và HH_D, gom theo TK cấp 1 (3–4 số)."""
    top, col = _snap_2tang(rows)
    if top is None:
        return []
    # CỘT MÃ TK DÒ THEO NHÃN, không gán cứng cột 0 (sửa 16/09/2026). Cùng mẫu S06-DN nhưng An
    # Taxi chèn một cột trống ở đầu: mã nằm ở cột B chứ không phải A. Bản cũ đọc `r[0]` nên ra
    # rỗng tuyệt đối — và rỗng ở đây KHÔNG gây lỗi, chỉ làm hai màn Thuế/Tồn kho trắng trơn.
    # Thêm alias "SHTK" (khối Dự án, 20/09/2026): cùng mẫu S06-DN nhưng cột mã viết tắt. Rơi xuống
    # mặc định 0 thì cột đó là ô trống -> hai màn Thuế/Tồn kho trắng trơn, đúng kiểu hỏng lặng lẽ
    # mà chú thích ngay trên đã cảnh báo.
    tk_j = next((j for j, c in enumerate(rows[top])
                 if "so hieu tai khoan" in _nd(c) or _nd(c) == "shtk"), 0)
    # PHÂN LOẠI TRƯỚC, CỘNG SAU — vì phải biết TK cha có mặt hay không rồi mới quyết cộng dòng nào.
    theo_key = {}   # (rt, dim1, dim2, key) -> [(tk, row), ...]
    chieu_cua = {}
    for r in rows[top + 2:]:
        tk = str(r[tk_j]).strip() if tk_j < len(r) and r[tk_j] not in (None, "") else ""
        if not tk[:3].isdigit():
            continue
        # TK con gộp về TK cha: '1331' -> '133', '333111' -> '3331', '15613' -> '156'.
        key = next((k for k in _THUE_TK if tk.startswith(k)), None)
        rt, dim1, dim2, chieu = ("THUE_D", *_THUE_TK[key], ) if key else (None, None, None, None)
        if key is None:
            key = next((k for k in _HH_TK if tk.startswith(k)), None)
            if key is not None:
                rt, dim1, dim2, chieu = "HH_D", _HH_TK[key], None, "no"
            elif congno_tong and tk in _CONGNO_TK:
                # KHỚP BẰNG ĐÚNG mã TK cấp 1, không theo tiền tố: CĐPS liệt kê cả 131 lẫn 1311/
                # 1316/1318, khớp tiền tố là cộng cha lẫn con.
                rt, chieu, dim1 = _CONGNO_TK[tk]
                dim2, key = None, tk
            else:
                continue
        k4 = (rt, dim1, dim2, key)
        theo_key.setdefault(k4, []).append((tk, r))
        chieu_cua[k4] = chieu

    gom = {}   # (rt, dim1, dim2, key) -> [dau, tang, giam, cuoi]
    for k4, ds_dong in theo_key.items():
        # CÓ DÒNG TK CHA THÌ CHỈ LẤY DÒNG ĐÓ, BỎ HẾT TK CON (khối Dự án, 20/09/2026).
        #
        # Bảng cân đối phát sinh của mẫu này liệt kê CẢ HAI cấp: sheet `CDSPS` vừa có dòng `133`
        # (3.670.571.141) vừa có dòng `1331003` (cũng 3.670.571.141), vừa có `152` vừa có
        # `1521003.VT`. Bản cũ cộng tuốt vì mọi TK con đều `startswith` khoá cha -> thuế GTGT được
        # khấu trừ ra ĐÚNG GẤP ĐÔI số của CĐKT (7.341.142.283 so với 3.670.571.141), tồn kho cũng
        # phồng tương tự. Số vẫn "có", chỉ là sai gấp đôi — kiểu hỏng không ai nhìn ra trên màn.
        #
        # Vẫn giữ đường cộng TK con cho mẫu CHỈ có cấp con (5 đơn vị đang chạy rơi vào nhánh này,
        # nên hành vi của chúng không đổi một ly).
        chinh = [x for x in ds_dong if x[0] == k4[3]]
        for _tk, r in (chinh or ds_dong):
            g = lambda k, _r=r: (_r[col[k]] if col[k] < len(_r)                       # noqa: E731
                                 and isinstance(_r[col[k]], (int, float)) else 0)
            pos, neg = ("no", "co") if chieu_cua[k4] == "no" else ("co", "no")
            acc = gom.setdefault(k4, [0.0, 0.0, 0.0, 0.0])
            acc[0] += g(f"dk_{pos}") - g(f"dk_{neg}")
            acc[1] += g(f"ps_{pos}")
            acc[2] += g(f"ps_{neg}")
            acc[3] += g(f"ck_{pos}") - g(f"ck_{neg}")
    facts = []
    for (rt, dim1, dim2, tk), (dau, tang, giam, cuoi) in gom.items():
        if not (dau or tang or giam or cuoi):
            continue
        pl = {"du_dau": round(dau * 1e-9, 9), "unit": "ty", "grain": "day"}
        if rt == "HH_D":
            pl.update({"tk": tk, "nhap": round(tang * 1e-9, 9), "xuat": round(giam * 1e-9, 9),
                       "cham_lc": 0.0, "ton_3m": 0.0, "ton_36": 0.0})
        else:
            pl.update({"ps_tang": round(tang * 1e-9, 9), "ps_giam": round(giam * 1e-9, 9)})
        facts.append((None, rt, dim1, dim1, cuoi, dim2, json.dumps(pl, ensure_ascii=False)))
    return facts


def _snap_sodu_facts(wb, ten_sheet, mau_cdkt=None, chi_so_du=False):
    """workbook -> [fact] CẢ CỤM SỐ DƯ (công nợ phải thu/trả, CĐKT, TS-NV, bảng cân đối phát sinh).

    Một chỗ duy nhất biết sheet nào nuôi report_type nào — trước đây danh sách này nằm rải trong
    `_snap_doc_ky`, thêm đơn vị thứ hai là phải chép đôi.
    """
    co = set(wb.sheetnames or [])

    def rows(sn):
        return [list(r) for r in wb[sn].iter_rows(values_only=True)]

    out = []
    for khoa, rt, chieu in (("pthu", "PTHU_D", "no"), ("ptra", "PTRA_D", "co")):
        sn = ten_sheet.get(khoa)
        if sn in co:
            out += _snap_congno_facts(rows(sn), rt, chieu)
    # Không có sổ công nợ chi tiết -> lấy TỔNG theo TK 131/331 từ CĐPS. Thẻ tổng của màn Công nợ
    # vẫn đúng, chỉ thiếu bảng Top 10 — trung thực hơn là để trắng cả màn.
    congno_tong = not (ten_sheet.get("pthu") or ten_sheet.get("ptra"))
    # "cdkt" xuất hiện HAI LẦN, cố ý: `_snap_cdkt_facts` lấy 3 mã tổng + TSCĐ cho thẻ KPI, còn
    # `_snap_tsnv_facts` lấy TOÀN BỘ dòng cho bảng cơ cấu Tài sản-Nguồn vốn. Hai report_type khác
    # nhau nên không cộng đôi.
    for khoa, boc in (("cdkt", lambda r: _snap_cdkt_facts(r, mau_cdkt)),
                      ("cdkt", lambda r: _snap_tsnv_facts(r, mau_cdkt)),
                      ("cdps", lambda r: _snap_cdps_facts(r, congno_tong))):
        sn = ten_sheet.get(khoa)
        if sn in co:
            out += boc(rows(sn))
    return _bo_phat_sinh(out) if chi_so_du else out


# Các khoá PHÁT SINH trong payload — bị ép về 0 khi nguồn là bản LUỸ KẾ (xem `_bo_phat_sinh`).
_PL_PHAT_SINH = ("ps_tang", "ps_giam", "nhap", "xuat")


def _bo_phat_sinh(facts):
    """Giữ SỐ DƯ, ép mọi cột PHÁT SINH về 0 — dùng cho bản luỹ kế (19/09/2026).

    Bản "Từ 01/09 Đến 15/09" của 2 HTX Xanh có cột số dư cuối kỳ ĐÚNG là số dư cuối ngày 15 (số dư
    là trạng thái tại một thời điểm, không phụ thuộc kỳ dài hay ngắn), nhưng cột phát sinh thì phủ
    15 NGÀY. Bỏ cả file vì cột phát sinh là mất luôn ngày 15 ở màn Tài sản-Nguồn vốn; giữ nguyên
    cột phát sinh là gán biến động nửa tháng cho một ngày. Lấy số dư, vứt phát sinh.

    `du_dau` cũng bị ép 0: nó là đầu kỳ của KHOẢNG 01→15, không phải đầu ngày 15. Phía đọc tự suy
    đầu kỳ từ ngày sớm nhất trong khoảng người xem chọn (`_shared._rows_dau_ky_ngay`), nên để 0 an
    toàn hơn là đưa một con số của quãng thời gian khác.
    """
    ra = []
    for f in facts:
        pl = f[6]
        if pl:
            try:
                d = json.loads(pl)
            except Exception:
                ra.append(f)
                continue
            for k in _PL_PHAT_SINH:
                if k in d:
                    d[k] = 0.0
            if "du_dau" in d:
                d["du_dau"] = 0.0
            f = (*f[:6], json.dumps(d, ensure_ascii=False))
        ra.append(f)
    return ra


def _sodu_ngay_cua_wb(wb, ten_sheet, ten_file, o_ngay_tu_o=False, ngay_tu_ten_file=False):
    """Workbook này là ảnh chụp số dư của NGÀY NÀO -> (ngày, None) | (None, lý do loại).

    Tách riêng khỏi `_bcqt_per_day` (18/09/2026) để đơn vị KHÔNG ở chế độ snapshot cũng dùng được
    cùng một phép nhận — xem `_sodu_quet_thu_muc`. Hai vế của phép nhận:

    1. FILE TỰ KHAI "Từ ngày X Đến ngày Y" (An Taxi, HTX Xanh). ĐÒI ĐÚNG MỘT NGÀY: `tu != den`
       nghĩa là bộ số dư đang ở khuôn luỹ kế, lúc đó "số dư cuối kỳ" không còn là số dư của MỘT
       ngày và gán cho `den` là bịa. Ca thật kỳ 2026-09: file `.D.20260915` của cả hai HTX khai
       01/09..15/09, file `.D.20260916` khai 16/09..16/09 — cùng thư mục, cùng tên mẫu, khuôn
       đổi giữa chừng y như Trạm sạc hồi 11/09.

    2. FILE CHỈ GHI MỘT Ô NGÀY (Xanh Vĩnh Phúc: `CĐKT` không có dòng "Từ ngày…", chỉ một ô ngày
       ngay dưới tiêu đề "BẢNG CÂN ĐỐI KẾ TOÁN"). Nhánh này PHẢI khai `ky_sodu_o_ngay` mới bật —
       không mở mặc định vì "thấy một ô ngày thì coi là kỳ" là kiểu suy diễn đã làm hỏng số một
       lần. Hai chốt chặn: trong 12 dòng đầu chỉ được có ĐÚNG MỘT ngày (nhiều ngày = không biết ô
       nào là kỳ), và ngày đó phải TRÙNG 8 số trong tên file — hai nguồn độc lập cùng nói một
       ngày thì mới nhận. Lệch nhau là loại và nói ra, không chọn bên nào.
    """
    sn = ten_sheet.get("cdkt") or ten_sheet.get("cdps")
    if sn not in (wb.sheetnames or ()):
        return None, f"không có sheet '{sn}'"
    rows = [list(r) for r in wb[sn].iter_rows(max_row=14, values_only=True)]
    ky = _bcqt_ky(rows)
    if ky:
        if ky[0] == ky[1]:
            return ky[0], None
        # BẢN LUỸ KẾ VẪN DÙNG ĐƯỢC PHẦN SỐ DƯ (19/09/2026). "Từ 01/09 Đến 15/09" nói rằng cột số
        # dư cuối kỳ là số dư CUỐI NGÀY 15 — đúng thứ cụm số dư cần; chỉ cột phát sinh là của 15
        # ngày. Trước đây loại cả file nên 2 HTX Xanh mất hẳn ngày 15 ở màn Tài sản-Nguồn vốn dù
        # số ngày đó có sẵn và CÂN tuyệt đối. Nay nhận, kèm cờ để `_snap_sodu_facts` vứt phát sinh.
        # Vẫn đòi `den` nằm trong kỳ và `tu <= den` — sai phạm vi thì không đoán.
        if ky[0] <= ky[1]:
            return ky[1], "luy_ke"
        return None, f"'{sn}' khai {ky[0]}..{ky[1]}, khoảng ngày không hợp lệ"
    # 3. FILE KHAI MỘT NGÀY BẰNG CHỮ: "Tại ngày 16 tháng 9 năm 2026" / "Ngày 16 tháng 9 năm 2026"
    #    (An Khách sạn). Nhận vô điều kiện vì câu này CHỈ nói được đúng một ngày — khác nhánh 2
    #    (một ô ngày trơ trọi, có thể là ngày in/ngày xuất nên phải khai cờ mới bật).
    #    Nhánh 1 chạy TRƯỚC nên file luỹ kế ("Từ ngày 01/09/2026 đến ngày 15/09/2026") đã bị loại,
    #    không lọt xuống đây. Kiểm chứng 18/09/2026: `CDPS` ngày 16 của An KS có số dư đầu kỳ khớp
    #    cuối kỳ ngày 15 ở 65/65 tài khoản -> đúng là ảnh chụp một ngày.
    for r in rows:
        for c in (r or ()):
            if not isinstance(c, str):
                continue
            m = re.search(r"[Nn]gày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})", c)
            if m:
                d, mm, y = (int(x) for x in m.groups())
                try:
                    return datetime.date(y, mm, d).isoformat(), None
                except ValueError:
                    return None, f"'{sn}' ghi ngày không hợp lệ: {c[:40]}"
    # 4. FILE KHÔNG KHAI NGÀY Ở ĐÂU CẢ -> lấy 8 số trong TÊN FILE (khối Dự án, 20/09/2026).
    #    Bản BCTC của kế toán Dự án để TRỐNG ô kỳ: dòng 7 của CĐKT nguyên văn là
    #    "Tại ngày   tháng   năm " — không điền gì. Ba nhánh trên đều không bắt được nên mọi file
    #    bị loại với lý do "không khai 'Từ ngày...'", tức cả cụm số dư im lặng không lên số.
    #
    #    PHẢI KHAI CỜ MỚI BẬT, không mở mặc định: tin tên file là bỏ mất lớp đối chứng thứ hai mà
    #    nhánh 2 cố ý dựng lên ("hai nguồn độc lập cùng nói một ngày thì mới nhận"). Chỉ bật ở đơn
    #    vị đã ĐO được rằng mỗi file đúng là một ảnh chụp riêng: khối Dự án có tồn kho
    #    23.728.445.510 -> 23.441.719.595 -> 23.268.102.376 -> 23.187.613.859 qua 4 file
    #    15→18/09, tức số nhúc nhích theo đúng ngày ghi ở tên file.
    #
    #    An toàn hơn ở CĐKT so với sheet dòng chảy: số dư là đại lượng THỜI ĐIỂM, nên kể cả file
    #    có phủ một khoảng thì cột "cuối kỳ" vẫn là số dư tại ngày chốt — khác hẳn doanh thu/chi
    #    phí, nơi gán số luỹ kế cho một ngày là phóng đại.
    if ngay_tu_ten_file:
        ngay_ten = _snap_day_of(ten_file)
        if ngay_ten:
            return ngay_ten, None
        return None, f"'{sn}' không khai kỳ và tên file cũng không có 8 số sau '.D.'"
    if not o_ngay_tu_o:
        return None, f"'{sn}' không khai 'Từ ngày ... Đến ngày ...'"
    ngay_o = {c.date().isoformat() if isinstance(c, datetime.datetime) else c.isoformat()
              for r in rows for c in (r or ())
              if isinstance(c, (datetime.datetime, datetime.date))}
    if len(ngay_o) != 1:
        return None, f"'{sn}' có {len(ngay_o)} ô ngày trong 12 dòng đầu, không chốt được kỳ"
    ngay = ngay_o.pop()
    ngay_ten = _snap_day_of(ten_file)
    if ngay_ten and ngay_ten != ngay:
        return None, f"'{sn}' ghi {ngay} nhưng tên file ghi {ngay_ten}"
    return ngay, None


def _sodu_quet_thu_muc(path, period, unit):
    """Quét MỌI file cùng kỳ trong thư mục -> ({ngày: [fact số dư]}, [file bị loại]).

    DÙNG CHO ĐƠN VỊ CÓ HAI HỌ FILE TÁCH RỜI (3 khối Xanh, 18/09/2026), khác Trạm sạc/An Taxi —
    ở đó P&L và số dư nằm trong CÙNG một workbook nên đọc kèm luôn (`_bcqt_per_day`,
    `_snap_doc_ky`), không cần quét thêm lượt nào:

        P&L    `B.6.XVP.D.202609.Baocaohqkdngay.xlsx`      — một file, sheet "01".."31"
        Số dư  `B.6.XVP.D.20260916.Baocaotaichinhrieng.xlsx` — mỗi file MỘT ảnh chụp cuối ngày

    Sheet nào nuôi report_type nào vẫn do `_snap_sodu_facts` quyết một chỗ duy nhất; ở đây chỉ lo
    chọn file và chọn ngày. Đơn vị khai THIẾU khoá nào trong `sheet_sodu` thì cụm đó không ra
    dòng — đó là cách thu hẹp phạm vi có chủ đích, không phải sót (xem `_UNITS`).
    """
    thu_muc = os.path.dirname(os.path.abspath(path))
    ten_sheet = unit.get("sheet_sodu") or {}
    o_ngay = bool(unit.get("ky_sodu_o_ngay"))
    ngay_ten_file = bool(unit.get("ngay_sodu_tu_ten_file"))
    gop, bo_qua, luy_ke = {}, [], []
    for ten in sorted(os.listdir(thu_muc)):
        if not ten.lower().endswith(".xlsx") or ten.startswith("~$"):
            continue
        # CHỈ HỌ ẢNH CHỤP (8 số sau `.D.`). File P&L cùng kỳ (6 số) cũng lọt `_period_of` nhưng
        # nó không có sheet số dư nào -> mở ra chỉ để đóng lại, và tệ hơn là đẻ một dòng "bị loại"
        # đọc như lỗi trong chẩn đoán mỗi lượt chạy.
        if _period_of(ten, True) != period or not _snap_day_of(ten):
            continue
        p = os.path.join(thu_muc, ten)
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        except Exception as e:
            bo_qua.append({"file": ten, "vi_sao": f"không mở được ({type(e).__name__})"})
            continue
        try:
            ngay, vi_sao = _sodu_ngay_cua_wb(wb, ten_sheet, ten, o_ngay, ngay_ten_file)
            if not ngay:
                bo_qua.append({"file": ten, "vi_sao": vi_sao})
                continue
            if ngay[:7] != period:
                bo_qua.append({"file": ten, "vi_sao": f"ảnh chụp {ngay}, ngoài kỳ {period}"})
                continue
            facts = _snap_sodu_facts(wb, ten_sheet, _MAU_CDKT.get(unit.get("mau_cdkt", "b01dn")),
                                     chi_so_du=(vi_sao == "luy_ke"))
            if vi_sao == "luy_ke":
                luy_ke.append(ngay)
            if not facts:
                bo_qua.append({"file": ten, "vi_sao": f"ảnh chụp {ngay} nhưng không bóc được "
                                                      f"dòng số dư nào"})
                continue
            # File sau ĐÈ file trước theo từng ngày, không cộng — số dư cùng một ngày mà cộng hai
            # bản xuất là gấp đôi. Cùng quy ước với `_bcqt_per_day`.
            gop[ngay] = facts
        finally:
            wb.close()
    return gop, bo_qua, sorted(set(luy_ke))


def _pl_quet_thu_muc(path, period, unit, da_co):
    """Quét họ file ảnh chụp để vét P&L cho những ngày file THÁNG chưa có -> ({ngày: [fact]}, chẩn đoán).

    LÝ DO TỒN TẠI (19/09/2026, 2 HTX Xanh): file tháng `B.6.HTX_*.D.202609.` dừng ở sheet "15" —
    kế toán ngừng xuất lại nó và chuyển số ngày sang chính file ảnh chụp, vốn từ 16/09 có thêm
    sheet "BC KQKD" (B02-HTX) khai đúng "Từ ngày 16/09 Đến ngày 16/09" và sheet "HQKD" cùng khuôn
    với sheet ngày của file tháng. Không có đường này thì 3 màn P&L đứng ở 15/09 trong khi nguồn
    đã có số tới 17/09, và không ai báo lỗi vì file tháng vẫn "đọc tốt" — nó chỉ thiếu sheet.

    HAI CHẶN BẮT BUỘC:
      · `tu == den` — theo LỜI KHAI của file, không suy từ tên (cùng luật `_tcode_snap_per_day`).
        Bản 15/09 khai "Từ 01/09 Đến 15/09" = LUỸ KẾ nửa tháng; nhận nó vào cụm dòng-chảy là cộng
        15 ngày đè lên 15 ngày đã có từ file tháng. Số dư thì ngược lại, luỹ kế vẫn dùng được —
        khác biệt đó là ranh giới flow/stock, xem `_bo_phat_sinh`.
      · `da_co` — ngày nào file THÁNG đã dựng được thì file tháng THẮNG, không ghi đè. Hai nguồn
        cùng ngày là cộng đôi (chúng vào chung `per_day` của cùng `source_file`).
    """
    cau_hinh = unit["pl_ho_file_rieng"]
    sheet_pl, sheet_ky = cau_hinh["sheet"], cau_hinh["sheet_ky"]
    thu_muc = os.path.dirname(os.path.abspath(path))
    gop, bo_qua = {}, []
    for ten in sorted(os.listdir(thu_muc)):
        if not ten.lower().endswith(".xlsx") or ten.startswith("~$"):
            continue
        if _period_of(ten, True) != period or not _snap_day_of(ten):
            continue
        try:
            wb = openpyxl.load_workbook(os.path.join(thu_muc, ten), data_only=True, read_only=True)
        except Exception as e:
            bo_qua.append({"file": ten, "vi_sao": f"không mở được ({type(e).__name__})"})
            continue
        try:
            if sheet_pl not in wb.sheetnames or sheet_ky not in wb.sheetnames:
                bo_qua.append({"file": ten, "vi_sao": f"không có sheet '{sheet_pl}'/'{sheet_ky}'"})
                continue
            ky = _bcqt_ky([list(r) for r in wb[sheet_ky].iter_rows(max_row=12, values_only=True)])
            if not ky:
                bo_qua.append({"file": ten, "vi_sao": f"sheet '{sheet_ky}' không có dòng "
                                                      f"'Từ ngày .. Đến ngày ..'"})
                continue
            if ky[0] != ky[1]:
                bo_qua.append({"file": ten, "vi_sao": f"luỹ kế {ky[0]}→{ky[1]}, không phải số "
                                                      f"riêng ngày — P&L bỏ (số dư vẫn dùng)"})
                continue
            ngay = ky[1]
            if ngay[:7] != period:
                bo_qua.append({"file": ten, "vi_sao": f"khai ngày {ngay}, ngoài kỳ {period}"})
                continue
            if ngay in da_co:
                bo_qua.append({"file": ten, "vi_sao": f"{ngay} đã có từ file tháng — nhường"})
                continue
            facts = _kqkd_facts([list(r) for r in wb[sheet_pl].iter_rows(values_only=True)])
            if not facts:
                bo_qua.append({"file": ten, "vi_sao": f"sheet '{sheet_pl}' không bóc được chỉ tiêu "
                                                      f"nào cho {ngay}"})
                continue
            gop[ngay] = facts
        finally:
            wb.close()
    return gop, bo_qua


def _snap_doc_ky(path, period, snap_sheet):
    """Đọc MỌI file báo cáo-theo-ngày cùng kỳ trong thư mục.

    -> ({den_ngay: (tu_ngay, {mã: (nhãn, giá trị)}, [fact số dư])}, chẩn đoán). Giá trị P&L là LUỸ
    KẾ hay RIÊNG NGÀY thì cứ xem `tu_ngay` — `_tcode_snap_per_day` xử tiếp. Fact số dư thì KHÔNG
    phụ thuộc chế độ đó: số dư cuối ngày luôn là số dư cuối ngày."""
    thu_muc = os.path.dirname(os.path.abspath(path))
    cum, bo_qua, lech_ten = {}, [], []
    for fn in sorted(os.listdir(thu_muc)):
        if not fn.lower().endswith((".xlsx", ".xlsm")) or fn.startswith("~$"):
            continue
        if _period_of(fn, True) != period or not _snap_day_of(fn):
            continue
        fp = os.path.join(thu_muc, fn)
        try:
            wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
        except Exception as e:
            # File hỏng/truyền dở (ca thật: `.D.20260829.` nặng đúng 7 byte, ruột là chuỗi
            # 'ban-moi' nhưng metadata vẫn ghi status ok) -> bỏ QUA file đó, KHÔNG bỏ cả kỳ.
            bo_qua.append({"file": fn, "vi_sao": f"không mở được ({type(e).__name__})"})
            continue
        try:
            if snap_sheet not in wb.sheetnames:
                bo_qua.append({"file": fn, "vi_sao": f"không có sheet '{snap_sheet}'"})
                continue
            byco = _tcode_byco([list(r) for r in wb[snap_sheet].iter_rows(values_only=True)])
            if not byco:
                bo_qua.append({"file": fn, "vi_sao": "không dò ra header 'Mã số'"})
                continue
            tu, ngay = _snap_ky(wb)
            sodu = _snap_sodu_facts(wb, _SHEET_SODU_MAC_DINH)
        finally:
            wb.close()
        if not ngay:
            bo_qua.append({"file": fn, "vi_sao": "không đọc được dòng 'Từ ngày .. Đến ngày ..' "
                                                 "-> không biết số là luỹ kế hay riêng ngày"})
            continue
        ten_ngay = _snap_day_of(fn)
        if ngay != ten_ngay:
            lech_ten.append({"file": fn, "ngay_trong_file": ngay, "ngay_ten_file": ten_ngay})
        cum[ngay] = (tu, byco, sodu)
    return cum, {**({"bo_qua_file": bo_qua} if bo_qua else {}),
                 **({"lech_ngay_ten_file": lech_ten} if lech_ten else {})}


def _tcode_snap_per_day(path, period, unit):
    """Bộ file báo cáo-theo-ngày của kỳ -> ([(ngay, facts)], chẩn đoán), số là PHÁT SINH RIÊNG NGÀY.

    Mỗi file TỰ KHAI kỳ của nó ở dòng "Từ ngày .. Đến ngày .." (xem `_snap_ky`) và hàm này theo
    đúng lời khai đó, KHÔNG suy từ tên file:
      · `tu == den`   -> file đã là số riêng ngày, dùng thẳng.
      · `tu <  den`   -> file là luỹ kế, số ngày = luỹ kế này − luỹ kế LUỸ KẾ liền trước.
    Trạm sạc đã đi qua cả hai: luỹ kế tới 10/09/2026, rồi xuất lại thành riêng ngày rạng sáng
    11/09. Hai chế độ lẫn trong cùng một kỳ thì trả cờ `tron_hai_che_do` — số vẫn dựng được nhưng
    đó là dấu hiệu nguồn đang chuyển khuôn, cần người nhìn."""
    cum, diag = _snap_doc_ky(path, period, unit["snapshot_sheet"])
    if not cum:
        return [], diag
    pp = unit.get("profit_pnlt", ("Lợi nhuận trước thuế",))
    ps = unit.get("pnlt_skip", ())
    ngays = sorted(cum)
    per_day, luyke_truoc, ghep, giam, che_do = [], None, [], [], set()
    truoc_ck, dut = {}, []        # chuỗi số dư: cuối ngày N vs đầu ngày N+1
    for ngay in ngays:
        tu, nay, sodu = cum[ngay]
        dau_thang = f"{ngay[:8]}01"
        if tu == ngay:                      # ── file khai đúng một ngày -> số đã là của ngày đó
            hieu = dict(nay)
        else:                               # ── file khai một dải -> luỹ kế, trừ mốc liền trước
            goc = luyke_truoc
            hieu = {c: (lab, (v or 0) - ((goc[1].get(c, (None, 0))[1] or 0) if goc else 0))
                    for c, (lab, v) in nay.items()}
            # THIẾU FILE GIỮA KỲ: hiệu này ôm luôn các ngày trống. Dồn vào ngày sau cùng là cách
            # duy nhất không bịa ra phân bổ, nhưng PHẢI nói ra để biểu đồ ngày đọc đúng.
            # Có mốc trước: hiệu phủ dải (mốc, ngay] -> gộp khi cách nhau hơn 1 ngày.
            # KHÔNG có mốc: giá trị thô phủ dải [mùng 1, ngay] -> gộp ngay khi `ngay` != mùng 1.
            # Hai vế phải viết riêng: dùng chung ngưỡng `> 1` thì ca "file luỹ kế đầu tiên của kỳ
            # là ngày 02" lọt lưới — số ngày 01 nằm gọn trong ngày 02 mà không cờ nào bật.
            moc = goc[0] if goc else dau_thang
            cach = (datetime.date.fromisoformat(ngay) - datetime.date.fromisoformat(moc)).days
            if (cach > 1) if goc else (ngay != dau_thang):
                ghep.append({"ngay": ngay, "gop_tu": moc,
                             "vi_sao": "kỳ thiếu file luỹ kế ở giữa" if goc else
                                       "kỳ thiếu file luỹ kế đầu tháng"})
        # MỐC TRỪ cho file luỹ kế kế tiếp = file nào BẮT ĐẦU TỪ MÙNG 1, vì giá trị thô của nó
        # chính là luỹ kế tới `ngay`. Gồm CẢ file ngày 01 (`tu == den == mùng 1`) — file đó vừa
        # đọc được là số riêng ngày vừa là luỹ kế tới mùng 1, hai cách cho cùng một con số.
        # BỎ SÓT CHỖ NÀY LÀ HỎNG CẢ NHÁNH LUỸ KẾ: ngày 01 rơi vào nhánh `tu == ngay` ở trên nên
        # không được ghi mốc, ngày 02 trừ vào rỗng và lấy nguyên số luỹ kế -> tổng kỳ bị cộng dư
        # đúng một lần ngày 01 (dựng dữ liệu giả 100/300/300 thì ra 400 thay vì 300).
        if tu == dau_thang:
            luyke_truoc = (ngay, nay)
        # PHÂN LOẠI CHẾ ĐỘ — chỉ kết luận khi chắc chắn. File ngày 01 có `tu == den == mùng 1`
        # KHÔNG nói lên điều gì (hai cách đọc trùng số), tính nó vào sẽ bật cờ `tron_hai_che_do`
        # oan cho một bộ file luỹ kế hoàn toàn bình thường.
        if tu != ngay:
            che_do.add("luy_ke")
        elif tu != dau_thang:
            che_do.add("rieng_ngay")
        # Số ngày ÂM = kế toán sửa số hồi tố. Giữ nguyên (đó là bút toán điều chỉnh thật, tự ý kẹp
        # về 0 là giấu mất điều chỉnh), chỉ báo ra.
        for c in ("T100", "T101", "T200"):
            if c in hieu and (hieu[c][1] or 0) < 0:
                giam.append({"ngay": ngay, "ma": c, "so_ngay": hieu[c][1]})
        facts = _tcode_facts_byco(hieu, pp, ps)
        # NGÀY CÓ FILE MÀ SỐ BẰNG 0 VẪN PHẢI CÓ DÒNG. `_tcode_facts_byco` bỏ qua mọi giá trị 0 —
        # đúng cho layout sheet-ngày (ở đó "không có dòng" = không có sheet = không biết), nhưng
        # ở đây ngày này CÓ FILE, nên 0 là một câu trả lời: ngày đó KHÔNG PHÁT SINH. Đã kiểm
        # chứng cho 01→05/09: BCĐPS ghi phát sinh = 0 ở mọi TK 511*, sổ nhật ký chỉ có bút toán
        # khấu hao. Thiếu dòng thì biểu đồ ngày đứt quãng, đọc thành "chưa có dữ liệu" — khác hẳn
        # "không phát sinh". Ghi thẳng 3 mã tổng = 0; tổng kỳ không đổi vì cộng thêm 0.
        # CHỈ zero-fill cho ngày CÓ FILE. Ngày thiếu file thì để trống và báo ở `ngay_thieu` —
        # bịa dòng 0 cho nó là khẳng định "không phát sinh" trong khi sự thật là "không biết".
        if facts:
            co = {f[2] for f in facts if f[1] == RT_HQKD}
            facts += [(None, RT_HQKD, ma, ma, 0.0)
                      for ma in (MA_DT, MA_CP, MA_LNTT) if ma not in co]
        else:
            facts = [(None, RT_HQKD, ma, ma, 0.0) for ma in (MA_DT, MA_CP, MA_LNTT)]
        # SỐ DƯ: gắn thẳng, KHÔNG đi qua `hieu`. Số dư cuối ngày đã là con số cần ghi cho ngày đó,
        # dù file P&L trong cùng workbook đang ở chế độ luỹ kế hay riêng ngày.
        facts += sodu
        # ĐỨT CHUỖI SỐ DƯ: số dư cuối ngày N phải bằng số dư đầu ngày N+1. Lệch = bộ file không
        # đồng nhất (một ngày nào đó chưa được xuất lại cùng lượt), và nó KHÔNG tự lộ ra thành lỗi
        # — màn Tồn kho vẫn vẽ đẹp, chỉ có đầu kỳ ≠ cuối kỳ − xuất. Ca thật Trạm sạc: file 01/09
        # lệch 174,8tr tồn kho · 526,3tr tổng tài sản · 1,61 tỷ trên BCĐPS so với ngày 02/09.
        # CHỈ soát loại có `du_dau` trong payload. BS_D (payload rỗng) và TS_D (chỉ có nguyên giá /
        # hao mòn) không mang đầu kỳ -> `dk` luôn 0 và mọi ngày đều bị báo đứt oan.
        for rt in {f[1] for f in sodu if f[6] and "du_dau" in f[6]}:
            ck = sum(f[4] for f in sodu if f[1] == rt)
            dk = sum((json.loads(f[6]).get("du_dau") or 0) * 1e9 for f in sodu if f[1] == rt and f[6])
            if truoc_ck.get(rt) is not None and abs(dk - truoc_ck[rt]) > 1:
                dut.append({"ngay": ngay, "loai": rt, "dau_ky": round(dk),
                            "cuoi_ky_hom_truoc": round(truoc_ck[rt]), "lech": round(dk - truoc_ck[rt])})
            truoc_ck[rt] = ck
        per_day.append((ngay, facts))
    # THIẾU FILE NGÀY — hậu quả KHÁC NHAU ở hai chế độ, nên phải báo riêng. Ở chế độ luỹ kế, ngày
    # trống được ngày sau ôm trọn (`ngay_ghep`) nên tổng kỳ vẫn đủ. Ở chế độ riêng ngày thì số của
    # ngày đó BIẾN MẤT KHỎI TỔNG THÁNG mà không có dấu hiệu gì — đây mới là ca âm thầm nguy hiểm.
    d0, d1 = (datetime.date.fromisoformat(x) for x in (ngays[0], ngays[-1]))
    thieu = [d.isoformat() for d in (d0 + datetime.timedelta(days=i)
                                     for i in range((d1 - d0).days + 1))
             if d.isoformat() not in cum]
    diag["nguon_ngay"] = {"so_file": len(ngays), "dau": ngays[0], "cuoi": ngays[-1],
                          "che_do": "+".join(sorted(che_do))}
    if thieu:
        diag["ngay_thieu"] = thieu
    if len(che_do) > 1:
        # Nguồn đang chuyển khuôn giữa kỳ. Số vẫn dựng đúng theo lời khai từng file, nhưng đây là
        # lúc dễ sai nhất nên phải để lại dấu cho người đọc log.
        diag["tron_hai_che_do"] = True
    if ghep:
        diag["ngay_ghep"] = ghep
    if giam:
        diag["so_ngay_am"] = giam
    if dut:
        diag["so_du_dut_chuoi"] = dut
    return per_day, diag


# ---------------------------------------------------------------------------------------------
# Layout "ht" — XE TẢI HƯNG THỊNH (HUNGTHINH/baocaohqkdngay/B5.HT.TCTC.M.<YYYYMM>.baocaongay.xlsx)
# ---------------------------------------------------------------------------------------------
# Mỗi ngày 1 sheet "{dd}.{mm}" ("01.08".."04.08" — spec user viết "sheet 01, 02, 03..." nên
# `_ht_day_sheets` nhận CẢ dạng số trần "01".."31"). Ô lấy số theo spec user 2026-08-06 (đã đối
# chiếu file thật T08, cột E = "Ngày {dd}" = tổng của F "Hưng Thịnh" + G "Thịnh Cường"):
#   Doanh thu HH, DV = E8+E9+E10+E11   ·   Tổng chi phí = E14   ·   Giá vốn = E15
#   CP tài chính = E21   ·   CP vận hành = E22   ·   LNTT = E41   ·   LNST = E43
# KHÔNG hardcode số dòng: dò theo MÃ ở cột mã số (T101.x/T200/T201/T202/T203/T300), vì mã ổn định
# còn thứ tự dòng đổi theo tháng (T203.1..T203.18 co giãn theo khoản mục kế toán phát sinh).
# ⚠ 3 điểm layout HT KHÁC hẳn "tcode" (GA/Trạm sạc) nên KHÔNG dùng lại `_tcode_facts` được:
#   1. Header cột mã số ghi "v" (không phải "Mã số") -> `_tcode_facts` không tìm ra header, chết
#      ngay bước đầu. Ở đây dò cột mã theo NỘI DUNG (cột nào chứa T100/T200/T300).
#   2. KHÔNG có mã "T101" tổng — doanh thu chỉ nằm ở 4 mã con T101.1..T101.4 (đúng spec E8..E11);
#      `_tcode_facts` đòi có "T101" nên sẽ trả [] cho mọi sheet.
#   3. LNST (E43) là dòng KHÔNG CÓ MÃ -> neo theo nhãn "Lợi nhuận sau thuế" ở cột Chỉ tiêu.
# dim1 giữ ĐÚNG tên bản THÁNG của HT đang dùng (verify raw_rows: 'Doanh thu HH, DV' / 'Giá vốn
# hàng bán' / 'Chi phí tài chính' / 'Chi phí vận hành' / 'Lợi nhuận sau thuế'; HQKD dùng mã
# 1000/1047/1112) để dòng NGÀY và dòng THÁNG khớp nhãn trên cùng một bảng.
# KHÔNG dựng "Lợi nhuận gộp"/các tỷ lệ: spec ghi rõ chúng "Tính toán trên Dashboard".
# Chuẩn hoá nhãn Y HỆT `agent_cli._derive_kqkd_tseries._canon` (2 mã, lý do xem trong `_ht_facts`).
_HT_CANON = {"T201": "Giá vốn hàng bán", "T103": "Thu nhập khác"}


def _ht_day_sheets(wb, period):
    """[(sheet_name, 'YYYY-MM-DD')] — sheet "{dd}.{mm}" (khớp ĐÚNG tháng của kỳ) hoặc "{dd}" trần.
    DÙNG CHUNG cho layout "ht" và "xdv" (cả 2 spec đều ghi "sheet 01, 02, 03... tương đương số ngày
    trong tháng", file thật đặt "01.08".."04.08")."""
    y, mm = int(period[:4]), int(period[5:7])
    out = []
    for s in wb.sheetnames:
        t = s.strip()
        m = re.fullmatch(r"(\d{1,2})(?:\.(\d{1,2}))?", t)
        if not m or not 1 <= int(m.group(1)) <= 31:
            continue
        if m.group(2) and int(m.group(2)) != mm:      # "05.07" = ngày 5 THÁNG 7 -> khác kỳ, bỏ
            continue
        out.append((s, f"{y:04d}-{mm:02d}-{int(m.group(1)):02d}"))
    return sorted(out, key=lambda x: x[1])


def _ht_facts(rows):
    """rows -> [(None, report_type, dim1, dim3, value_VND)] cho 1 ngày. [] nếu sai layout."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    ten_j = next(j for j, c in enumerate(hdr) if _nd(c) == "chi tieu")
    # Cột giá trị: header "Ngày {dd}"; fallback = cột ngay sau "Chỉ tiêu" (cột E ở file mẫu).
    val_j = next((j for j, c in enumerate(hdr) if j > ten_j and re.fullmatch(r"ngay ?\d{0,2}", _nd(c))),
                 ten_j + 1)
    # Cột mã số: header ghi "v" (vô nghĩa) -> tìm cột NÀO chứa các mã T-series trong phần dữ liệu.
    body = rows[hdr_i + 1:]
    ma_j = next((j for j in range(0, ten_j)
                 if sum(1 for r in body if j < len(r) and re.fullmatch(r"T\d{3}(\.\d+)?",
                                                                      str(r[j] or "").strip())) >= 3), None)
    if ma_j is None:
        return []

    # byco: mã -> (nhãn gốc, giá trị). Nhận CẢ mã con "T203.7" (bản tháng cũng gom vậy).
    byco, lnst = {}, None
    for r in body:
        v = _num(r[val_j]) if val_j < len(r) else None
        code = str(r[ma_j] or "").strip() if ma_j < len(r) else ""
        if re.fullmatch(r"T\d{3}(\.\d+)?", code) and code not in byco:
            lab = str(r[ten_j]).strip() if ten_j < len(r) and r[ten_j] not in (None, "") else code
            byco[code] = (lab, v)
        if lnst is None and ten_j < len(r) and _nd(r[ten_j]).startswith("loi nhuan sau thue"):
            lnst = v
    if "T200" not in byco or "T300" not in byco:
        return []

    def val(code):
        return byco.get(code, (None, None))[1]

    facts = []
    # "Doanh thu thuần" = Σ T101.x (hoặc T101 gộp nếu có), KHÔNG phải T100 — Y HỆT bản THÁNG
    # (agent_cli._derive_kqkd_tseries, chốt user 2026-07-19 đối chiếu file gốc T05: T100 = ΣT101.x +
    # T102 DT tài chính + T103 thu nhập khác, nên T100 THỪA cho chỉ tiêu "doanh thu thuần"). Sheet
    # NGÀY của HT không có dòng T101 gộp, chỉ có T101.1..T101.4 = đúng "E8+E9+E10+E11" của spec.
    t101 = [v for c, (_, v) in byco.items() if re.fullmatch(r"T101\.\d+", c) and v is not None]
    dt = val("T101") if val("T101") is not None else (sum(t101) if t101 else val("T100"))
    if dt:
        facts.append((None, RT_HQKD, MA_DT, MA_DT, dt))
        facts.append((None, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
        facts.append((None, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
    if val("T200"):
        facts.append((None, RT_HQKD, MA_CP, MA_CP, val("T200")))
    if val("T300"):
        facts.append((None, RT_HQKD, MA_LNTT, MA_LNTT, val("T300")))
    # LNST: bản THÁNG gán = T300 vì sheet tháng KHÔNG có dòng LNST/thuế TNDN riêng. Sheet NGÀY thì CÓ
    # (dòng "LỢI NHUẬN SAU THUẾ TNDN" = E43, ngay dưới dòng "Thuế TNDN") -> ưu tiên dòng thật, chỉ
    # rơi về T300 khi thiếu. Hôm nay thuế = 0 nên 2 cách ra cùng số; khi HT phát sinh thuế thật thì
    # dòng thật MỚI đúng, còn T300 sẽ là trước thuế.
    lnst = lnst if lnst is not None else val("T300")
    if lnst:
        facts.append((None, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lnst))
    # Lợi nhuận gộp = DOANH THU THUẦN (`dt` = Σ T101.x, tính ở trên) − giá vốn T201.
    # SỬA 11/08/2026 — trước lấy T100, tức cộng luôn T102 DT tài chính + T103 thu nhập khác vào lãi
    # gộp. Chú thích cũ ghi "bản tháng dùng T100 ở ĐÚNG chỗ này" là MÔ TẢ SAI: bản tháng cũng sai y
    # hệt và đã sửa cùng lượt (agent_cli._derive_kqkd_tseries). Đo kỳ 06/2026: HT 4,535576 -> đúng
    # 4,189939 tỷ, thổi lên 345,6 triệu.
    if val("T201") is not None and dt is not None and (dt - val("T201")):
        facts.append((None, RT_PNLT, "Lợi nhuận gộp", "Lợi nhuận gộp", dt - val("T201")))
    # MỌI mã còn lại -> PNLT giữ NHÃN GỐC (T101.x dòng xe · T102 · T201.x giá vốn chi tiết · T202 ·
    # T203 + T203.x 18 khoản chi phí vận hành). Chuẩn hoá 2 nhãn y hệt bản tháng: T201 -> 'Giá vốn
    # hàng bán' (nguồn HT gõ 'Gía vốn' — dấu sắc trên i — làm metrics.build_revenue lọc
    # ILIKE '%giá vốn%' ACCENT-SENSITIVE trượt) và T103 -> 'Thu nhập khác' (HT gõ 'Doanh thu khác').
    # Mã CON giữ nguyên nhãn typo -> KHÔNG khớp filter -> KHÔNG đếm đôi với dòng tổng đã chuẩn hoá.
    for code, (lab, v) in byco.items():
        if code in ("T100", "T200", "T300") or not v:
            continue
        ten = _HT_CANON.get(code, lab)
        facts.append((None, RT_PNLT, ten, ten, v))
    # 02_CHIPHI: các mã CON TRỰC TIẾP của T200 = '^T2\d{2}$' trừ T200 (T201 giá vốn · T202 CP tài
    # chính · T203 CP vận hành; T2xx.y là con của chúng -> loại, tránh cộng trùng). Chỉ dựng khi Σ
    # con PHỦ HẾT T200 (sai số 1%) — y hệt điều kiện `_covers` của bản tháng, tránh breakdown thiếu.
    cp = [(c, byco[c][0], byco[c][1]) for c in byco
          if re.fullmatch(r"T2\d{2}", c) and c != "T200" and byco[c][1]]
    if cp and val("T200") and abs(sum(x[2] for x in cp) - val("T200")) <= abs(val("T200")) * 0.01:
        for _c, nhom, v in cp:
            facts.append((None, RT_CHIPHI, nhom, nhom, v))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "xdv" — XƯỞNG DỊCH VỤ VINFAST (XDV/baocaohqkdngay/B.2.TC.TCKT.M.<YYYYMM>.BaocaoHQKD.xlsx)
# ---------------------------------------------------------------------------------------------
# Mỗi ngày 1 sheet "{dd}.{mm}". Header dòng 8: A "Mã số" · B "Chỉ tiêu" · C "Kỳ này" (TỔNG) ·
# D..H (TK nợ/TK có/Mã phí/Mã NS/Công thức, BỎ QUA) · I→V = 14 cost center.
# Mã B-series (spec user 2026-08-06): B100 DT · B210 DT thuần · B300 giá vốn · B410 lãi gộp ·
# B500 tổng CP xưởng (= B600 nhân sự + B700 hoạt động) · B810 CP cố định · B822 lãi vay ·
# B833 CP khác · B840 LNTT · B900 LNST.
#
# ⚠ LỆCH GIỮA 2 CỘT CỦA CHÍNH SPEC — "Tổng chi phí": cột "Cách lấy (nguyên văn)" ghi
# `B500+B810+B822+B833` (KHÔNG có giá vốn), nhưng cột "Công thức" ngay bên cạnh lại ghi
# "Tổng CP = **Giá vốn** + CP QLDN + CP bán hàng + CP khác + CP tài chính" (CÓ giá vốn). Bản THÁNG
# của XDV (agent_cli `_kqkd_recs_xdv`/`_chiphi_recs_xdv`) đã chốt 2026-07-17 là **GỒM B300**, có
# audit đối soát BCHN Tập đoàn (37,71 khớp hợp nhất) và để `byNhom == byKhoi`. Deriver NGÀY theo
# ĐÚNG bản tháng (1047 = B300+B500+B810+B822+B833): nếu bỏ B300 thì cùng thẻ "Chi phí" sẽ mang 2
# nghĩa khác nhau khi bấm qua lại Tháng/Ngày (ngày ~0,2 tỷ vs tháng ~13,3 tỷ) — sai nghiêm trọng
# hơn nhiều so với việc lệch chữ với cột "nguyên văn".
#
# Mã cost center lấy Y HỆT bản tháng (`agent_cli._XDV_BRANCH_CC`, khớp trust_me_bro.xlsx) — dò theo
# TỪ KHOÁ trong tên cột, không theo vị trí. Cả 14 đều pháp nhân TC (không có case chéo pháp nhân
# như SRVF UB_SR) nên không cần `_CC_CONGTY`. Verify file T08: 14/14 cột nhận diện được và
# **Σ 14 CC = ĐÚNG cột C** ở cả 10 mã × 3 ngày -> ghi THEO CC (không ghi thêm dòng tổng, tránh
# đếm đôi; tổng khối = Σ cost center, đúng nhánh cc_v của repository._per_file_resolved).
_XDV_CC = [
    ("ocean park", "OCP_XDV"), ("long bien", "LB_XDV"), ("smart city", "SMC_XDV"),
    ("ha long", "HL_XDV"), ("cam pha", "CP_XDV"), ("xuan mai", "XM_XDV"),
    ("uong bi", "UB_XDV"), ("tuyen quang", "TQ_XDV"), ("vinh phuc", "VP_XDV"),
    ("son tay", "ST_XDV"), ("dai tu", "ĐT_XDV"), ("viet tri", "VT_XDV"),
    ("ha khanh", "HK_XDV"), ("ho chi minh", "HCM_XDV"),
]
_XDV_CP_CODES = ("B300", "B500", "B810", "B822", "B833")     # cấu thành 1047 (GỒM giá vốn B300)


def _xdv_facts(rows):
    """rows -> [(cost_center, report_type, dim1, dim3, value_VND[, dim2])] cho 1 ngày. [] nếu sai
    layout. Phần tử thứ 6 (dim2 = yếu tố chi phí) CHỈ có ở dòng CHIPHI — xem cuối hàm."""
    hdr_i = next((i for i, r in enumerate(rows[:14])
                  if any(_nd(c) == "ma so" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    ma_j = next(j for j, c in enumerate(hdr) if _nd(c) == "ma so")
    ten_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("chi tieu")), ma_j + 1)
    cc_cols = [(cc, j) for kw, cc in _XDV_CC
               for j in [next((k for k, c in enumerate(hdr) if isinstance(c, str) and kw in _nd(c)),
                              None)] if j is not None]
    if len(cc_cols) < 10:        # thiếu chi nhánh = đổi layout -> báo sai layout, KHÔNG ghi số lệch
        return []

    byco = {}
    for r in rows[hdr_i + 1:]:
        c = str(r[ma_j] or "").strip() if ma_j < len(r) else ""
        if re.fullmatch(r"B\d{3}", c) and c not in byco:
            byco[c] = r
    if "B840" not in byco or "B300" not in byco:
        return []

    def v(code, j):
        r = byco.get(code)
        return _num(r[j]) if r is not None and j < len(r) else None

    def vsum(codes, j):
        xs = [v(c, j) for c in codes]
        return sum(x for x in xs if x) if any(x is not None for x in xs) else None

    facts = []
    for cc, j in cc_cols:
        dt = v("B210", j) if v("B210", j) is not None else v("B100", j)
        for rt, dim1, val in (
                (RT_HQKD, MA_DT, dt),
                (RT_DTHU, "Doanh thu thuần", dt),
                (RT_HQKD, MA_CP, vsum(_XDV_CP_CODES, j)),
                (RT_HQKD, MA_LNTT, v("B840", j)),
                (RT_PNLT, "Doanh thu thuần", dt),
                # "Doanh thu HH, DV" = KPI #1 của bảng 50 chỉ tiêu (overview.py dò
                # dim1_ilike "doanh thu hh%") và cột DT GỘP của màn Doanh thu (revenue.py
                # `_gross_of`). Bản NGÀY của XDV trước đây không ghi dòng này nên ô #1 ở
                # "CHỈ TIÊU TÀI CHÍNH" tab Ngày để trống, còn ô #2 "Doanh thu thuần" thì có
                # số — cạnh nhau mà một bên trắng.
                # LẤY B100 (doanh thu GỘP, trước giảm trừ B200) — ĐÚNG như bản THÁNG dùng cho
                # chỉ tiêu này, rơi về `dt` nếu sheet ngày thiếu B100. Bản đầu (03/09) ghi thẳng
                # `dt` cho nhanh; giá trị không đổi vì B200 = 0 ở **mọi** sheet ngày T08/2026 (đã
                # đếm: 0/32 ngày có giảm trừ), nhưng B100 mới là mã của chỉ tiêu — ngày nào XDV
                # thật sự có giảm trừ thì gộp phải khác thuần, không được bằng nhau vì code ép.
                # KHÔNG cộng đôi: hai dim1 khác nhau, mọi consumer PNLT đều lọc theo dim1.
                (RT_PNLT, "Doanh thu HH, DV", v("B100", j) if v("B100", j) is not None else dt),
                (RT_PNLT, "Giá vốn hàng bán", v("B300", j)),
                (RT_PNLT, "Lợi nhuận gộp", v("B410", j)),
                (RT_PNLT, "Lợi nhuận sau thuế", v("B900", j)),
                # HOẠT ĐỘNG KHÁC/TÀI CHÍNH — cho tab Ngày của màn Hiệu quả kinh doanh có số như
                # tab Tháng. B831 là TK 511124 (doanh thu chiến dịch), KHÔNG gộp vào thu nhập
                # khác — xem `agent_cli._derive_kqkd_xdv`, cùng quy ước.
                (RT_PNLT, "Doanh thu chiến dịch", v("B831", j)),
                (RT_PNLT, "Thu nhập khác", v("B832", j)),
                (RT_PNLT, "Doanh thu tài chính", v("B821", j))):
            if val:
                facts.append((cc, rt, dim1, dim1, val))
        # ---- CẤU TRÚC DOANH THU theo NGÀY (04/09/2026) ----
        # Không có mấy dòng này thì bảng "Cấu trúc Doanh thu" ở tab Ngày chỉ còn ĐÚNG HAI dòng
        # "Doanh thu bán hàng & CCDV" và "Doanh thu thuần", mà XDV lại không có giảm trừ nên hai
        # dòng đó bằng nhau y hệt — nhìn như bảng lỗi lặp dòng, trong khi phần chi tiết giải thích
        # con số thì file NGÀY có sẵn (B110-B140 nằm ở cả 32/32 sheet ngày, verify T08/2026).
        # B150 "Doanh thu Sửa chữa động cơ" CHƯA có ở bản ngày (0/32 sheet) -> tự khuyết, không ép.
        # Verify T08/2026: Σ B110..B140 = ĐÚNG BẰNG B100 từng ngày.
        for _ma, _ten in (("B110", "Doanh thu công việc (XHĐ)"),
                          ("B120", "Doanh thu phụ tùng (XHĐ)"),
                          ("B130", "Chiết khấu phụ tùng bảo hành (XHĐ)"),
                          ("B140", "Doanh thu cứu hộ 247"),
                          ("B150", "Doanh thu Sửa chữa động cơ")):
            _x = v(_ma, j)
            if _x:
                facts.append((cc, RT_PNLT, _ten, _ten, _x))
        # B200 ghi CẢ KHI = 0 (khác vòng trên dùng `if val`): "có khoản mục, hôm nay không phát
        # sinh" là thông tin khác hẳn "không có nguồn" — FE tự gom vào nhóm ẩn kèm nhãn đếm.
        _gt = v("B200", j)
        if _gt is not None:
            facts.append((cc, RT_PNLT, "Các khoản giảm trừ doanh thu",
                          "Các khoản giảm trừ doanh thu", _gt))
        # Cơ cấu chi phí (spec #10) — tách B500 -> B600 nhân sự + B700 hoạt động NẾU cộng khớp,
        # y hệt `agent_cli._chiphi_recs_xdv`; không khớp thì để nguyên B500.
        # dim2 = 6 YẾU TỐ CHI PHÍ của XDV (review DB 04/09/2026) — ÉP THEO MÃ, y hệt bản THÁNG
        # `agent_cli._chiphi_recs_xdv`. Trước đây bản NGÀY không ghi dim2 nào, nên biểu đồ "Yếu tố
        # chi phí" ở tab Ngày TRỐNG TRƠN trong khi tab Tháng có đủ 6 cột — cùng một màn, bấm qua
        # lại là mất số. dim1 (Nhóm) và dim3 (tên gốc) GIỮ NGUYÊN -> tổng theo nhóm không đổi một
        # đồng nào. Nhãn dùng lại ĐÚNG chuỗi của bản tháng để chú giải hai tab khớp nhau.
        _YT_KHAC_XDV = "CP hoạt động khác XDV"
        b500, b600, b700 = v("B500", j), v("B600", j), v("B700", j)
        cp_groups = [("B300", "Giá vốn hàng bán", "CP Giá vốn")]
        if None not in (b500, b600, b700) and abs((b600 + b700) - b500) < 1000:
            cp_groups += [("B600", "Chi phí nhân sự", "CP nhân sự"),
                          ("B700", "Chi phí hoạt động xưởng", _YT_KHAC_XDV)]
        else:
            cp_groups += [("B500", "Chi phí xưởng dịch vụ", _YT_KHAC_XDV)]
        # B810 TÁCH -> B811 (mặt bằng) + B812 (khấu hao): hai mã con mang HAI yếu tố KHÁC NHAU nên
        # không ép được một nhãn cho cả cụm. Guard Σcon ≈ cha (<1000đ) y như nhánh B500 ở trên; sai
        # thì giữ nguyên B810 lump với nhãn khấu hao (cấu phần lớn hơn) để tổng nhóm không hụt.
        b810, b811, b812 = v("B810", j), v("B811", j), v("B812", j)
        if None not in (b810, b811, b812) and abs((b811 + b812) - b810) < 1000:
            cp_groups += [("B811", "Chi phí cố định", "CP mặt bằng"),
                          ("B812", "Chi phí cố định", "CP khấu hao TSCĐ")]
        else:
            cp_groups += [("B810", "Chi phí cố định", "CP khấu hao TSCĐ")]
        cp_groups += [("B822", "Chi phí tài chính", "CP lãi vay"),
                      ("B833", "Chi phí khác", _YT_KHAC_XDV)]
        for code, nhom, yeuto in cp_groups:
            x = v(code, j)
            if x:
                ten = str(byco[code][ten_j]).strip() if ten_j < len(byco[code]) and byco[code][ten_j] else nhom
                facts.append((cc, RT_CHIPHI, nhom, ten, x, yeuto))
    return facts


# ---------------------------------------------------------------------------------------------
# Layout "duan" — Khối Dự án (DUAN/baocaohqkdngay), mỗi ngày 1 sheet "1".."31"
# ---------------------------------------------------------------------------------------------
# Cột cost center dò theo TỪ KHOÁ (chứa, không cần khớp hệt) vì nhãn cột đổi nhẹ theo tháng (vd
# "Yên Bình 3"). Mã CC lấy Y HỆT bản THÁNG (agent_cli._DA_PROJECT_CC) — kể cả quy ước NGƯỢC viết
# tắt Tân Thịnh<->Yên Bình đã xác nhận nguồn, xem docstring đầu file. "Bình phước" mới, chưa có
# trong master_data -> mã tự đặt (giống Núi Pháo/Quảng Ngãi bản tháng).
_CC_DUAN = [("cao bang", "CB_DA"), ("tan thinh", "TT_DA"), ("lang son", "LS_DA"),
            ("yen binh", "YB_DA"), ("phu quoc", "PQ_DA"), ("quang son", "QS_DA"),
            ("nui phao", "NUIPHAO_DA"), ("quang ngai", "QUANGNGAI_DA"), ("tho chu", "TC_DA"),
            ("binh phuoc", "BINHPHUOC_DA")]

# (khoá -> (nhãn chuẩn hoá, exact?)). Đa số EXACT (không startswith) vì nhãn ngắn dễ bị dòng con
# "nuốt" nhầm — vd "chi phi khac" (mục X.2, mã neo cp_khac) là PREFIX của "Chi phí khác tại dự
# án" (mục con 1.7 của Giá vốn IV, đứng TRƯỚC trong sheet) -> startswith sẽ khoá nhầm dòng đó.
# Chỉ lntt/lnst dùng startswith vì nhãn gốc có hậu tố đổi được "(EBT)"/"(EAT)".
# Khối I (DT gộp) và II (giảm trừ) CHỈ dùng làm ĐƯỜNG LÙI cho `dt_thuan`, không sinh chỉ tiêu
# riêng — xem `_duan_dt`. Neo EXACT để không dính các dòng con "Doanh thu thực hiện dự án"…
_DUAN_ANCHOR = {
    "dt_gross": ("doanh thu", True),              # I   (đường lùi)
    "giam_tru": ("cac khoan giam tru", True),     # II  (đường lùi)
    "dt_thuan": ("doanh thu thuan", True),        # III
    "gia_von": ("gia von hang ban", True),        # IV
    "ln_gop": ("lai gop", True),                  # V (nhãn gốc "Lãi gộp")
    "cp_bien_doi": ("chi phi bien doi", True),    # VI
    "cp_co_dinh": ("chi phi co dinh", True),      # VIII
    "cp_tai_chinh": ("chi phi tai chinh", True),  # IX.2
    "cp_khac": ("chi phi khac", True),            # X.2
    "lntt": ("loi nhuan truoc thue", False),      # XI "...(EBT)"
    "lnst": ("loi nhuan sau thue", False),        # XII "...(EAT)"
}
# Tổng chi phí = IV + VI + VIII + IX.2 + X.2 (đúng công thức "E25+E46+E74+E80+E83" của Mapping) —
# 5 nhóm Y HỆT ANTAXI (không có mục "phân bổ chung" cấp I riêng, đã lồng trong "Chi phí khác").
# 7 KHOẢN MỤC CON CỦA GIÁ VỐN (mục IV.1.1 -> IV.1.7 của sheet ngày) — nuôi khối "Cơ cấu giá vốn
# theo khoản mục" của màn `duan2`, đúng dòng 25 (GIÁ VỐN) tới dòng 32 của Mapping Dự án.
#
# VÌ SAO `RT_DUAN_GV` LÀ REPORT_TYPE RIÊNG, KHÔNG PHẢI THÊM `dim1` VÀO `CHIPHI_D`: giá vốn cấp I
# ("Giá vốn hàng bán", đã có trong `_DUAN_CP`) BẰNG TỔNG 7 dòng con này. Mọi màn đều cộng các dòng
# cùng report_type, nên nhét chung là ĐẾM ĐÔI toàn bộ giá vốn — đúng cái bẫy đã ghi ở `_bay` của
# `xdv_hqkd_ngay` ("dòng Khối + 14 dòng xưởng trong cùng một report_type").
#
# (từ khoá đã qua `_nd`, nhãn chuẩn). `_nd` ở file này GIỮ khoảng trắng và dấu hai chấm — từ khoá
# phải viết rời chữ ("nhan cong truc tiep"), đừng chép kiểu dính liền của `spec_extract._nd`.
# Từ khoá phải ĐỦ DÀI để không bắt nhầm nhau: "nhan cong truc tiep" chứ không phải "nhan cong",
# vì mục VI.1 "Chi phí nhân sự" và bảng lương đều có "nhân công".
_DUAN_GV_CT = [
    ("nvl", "Chi phí NVL"),
    ("nhan cong truc tiep", "Chi phí nhân công trực tiếp"),
    ("nhien lieu", "Chi phí nhiên liệu"),
    ("khau hao", "Chi phí khấu hao"),
    ("vtsc", "Chi phí VTSC"),
    ("thau phu", "Chi phí thầu phụ"),
    ("khac tai du an", "Chi phí khác tại dự án"),
]

_DUAN_CP = [
    ("gia_von", "Giá vốn hàng bán", "Giá vốn hàng bán"),
    ("cp_bien_doi", "Chi phí biến đổi", "Chi phí biến đổi"),
    ("cp_co_dinh", "Chi phí cố định", "Chi phí cố định"),
    ("cp_tai_chinh", "Chi phí tài chính", "Chi phí tài chính"),
    ("cp_khac", "Chi phí khác", "Chi phí khác"),
]


# Bản đồ BỔ SUNG do admin duyệt, nằm ở DB (`cost_center_map`, migration 0070) — xem
# `scripts/soat_cost_center.py`. Để deriver KHÔNG phải deploy mỗi lần kế toán thêm một dự án:
# admin bấm duyệt ở chuông 🔔 là lượt nạp kế tiếp đọc được ngay.
#
# CHỈ lấy `da_duyet`. Dòng `cho_duyet` cố ý KHÔNG dùng: nó mới là ĐỀ XUẤT của máy, mà mã cost
# center đi vào mọi báo cáo và mọi bộ lọc — đặt sai thì số chạy sang sai đơn vị, lần ngược rất
# khó. Thà cột nằm chờ (đúng hành vi hôm nay, đã có cảnh báo ở chuông) còn hơn lên nhầm mã.
_cc_db_cache = None


def _cc_da_duyet(layout: str = "duan"):
    """[(tu_khoa_chuan_hoa, ma_cost_center)] từ DB. Lỗi DB -> [] (giữ nguyên bản đồ trong code).

    Nuốt lỗi CÓ CHỦ Ý: bảng có thể chưa tồn tại (DB chưa chạy migration 0070) hoặc đang chạy
    ngoài luồng backend. Nạp số quan trọng hơn bản đồ bổ sung, và thiếu nó thì hành vi lùi về
    đúng như trước khi có tính năng này chứ không hỏng.
    """
    # Cache THEO LAYOUT, không phải một biến chung: từ 18/09 bộ soát gọi hàm này cho 4 layout
    # (duan/srvf/xdv/antaxi) trong cùng một tiến trình — cache chung sẽ trả bản đồ của layout gọi
    # ĐẦU TIÊN cho mọi layout sau, tức gán mã sai đơn vị. Bắt được ngay khi mở rộng, chưa kịp chạy.
    global _cc_db_cache
    if _cc_db_cache is None:
        _cc_db_cache = {}
    if layout in _cc_db_cache:
        return _cc_db_cache[layout]
    out = []
    try:
        with psycopg.connect(DB_URL) as conn:
            for ten, cc in conn.execute(
                    "SELECT ten_cot_chuan, cost_center FROM cost_center_map "
                    "WHERE layout=%s AND trang_thai='da_duyet'", (layout,)).fetchall():
                if ten and cc:
                    out.append((_nd(ten), cc))
    except Exception:                                    # noqa: BLE001 - xem docstring
        out = []
    _cc_db_cache[layout] = out
    return out


def _ma_cost_center(ten_cot: str):
    """Tên cột Excel -> mã cost center. None = CHƯA CÓ MÃ (chờ admin duyệt).

    Bản đồ trong code đi TRƯỚC bản đồ DB: 10 dự án gốc đã được đối chiếu tay với master_data,
    không để một dòng duyệt nhầm trong DB ghi đè chúng.
    """
    n = _nd(ten_cot)
    return (next((cc for kw, cc in _CC_DUAN if kw in n), None)
            or next((cc for kw, cc in _cc_da_duyet() if kw and kw in n), None))


def _duan_facts(rows):
    """rows -> [(cost_center, report_type, dim1, dim3, value_VND)]. [] nếu sai layout.

    KHÁC mọi layout khác: nhãn cột "CHỈ TIÊU" (dòng "STT|CHỈ TIÊU|DỰ ÁN") và tên cost center
    ("Tổng Dự án|HO Dự án|<từng dự án>") nằm ở 2 DÒNG KHÁC NHAU (dòng sau) — `_kqkd_scan` giả
    định cùng 1 dòng nên không dùng chung được, phải dò riêng."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "tong du an" for c in r if c is not None)), None)
    ten_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None or ten_i is None:
        return []
    hdr = rows[hdr_i]
    ten_j = next(j for j, c in enumerate(rows[ten_i]) if _nd(c) == "chi tieu")
    # BỎ cột "Tổng Dự án"/"HO Dự án" — chỉ lấy 7 cột dự án (E:K, đúng theo Mapping) để Σ cost
    # center không đếm đôi với cột tổng (verify 01/08: Cao Bằng 392.421.525 + Phú Quốc
    # 239.351.852 = 631.773.377 = đúng "Tổng Dự án"; "HO Dự án" luôn 0 các ngày đã verify).
    cols = [(cc, j) for j, c in enumerate(hdr) if isinstance(c, str)
            for cc in [_ma_cost_center(c)] if cc]
    if len(cols) < 3:
        return []

    anchored, gv_ct = {}, {}
    for r in rows[ten_i + 1:]:
        if not r or ten_j >= len(r):
            continue
        n = _nd(r[ten_j])
        for key, (pref, exact) in _DUAN_ANCHOR.items():
            if key not in anchored and (n == pref if exact else n.startswith(pref)):
                anchored[key] = r
        # 7 dòng con của IV.1, nhãn "Giá vốn: Chi phí ..." — dò theo TỪ KHOÁ chứ không khớp hệt
        # vì nhãn đổi nhẹ theo tháng. Cổng `startswith("gia von")` giữ cho từ khoá ngắn không bắt
        # nhầm dòng khác: "thau phu" cũng nằm trong "Doanh thu bán dầu thầu phụ" (khối I) và
        # "khau hao" nằm trong "CP Khấu hao và phân bổ" (khối VIII).
        if n.startswith("gia von"):
            for kw, _ten in _DUAN_GV_CT:
                if kw in n and kw not in gv_ct:
                    gv_ct[kw] = r
    if "dt_thuan" not in anchored:
        return []

    def val(key, j):
        r = anchored.get(key)
        return _num(r[j]) if r is not None and j < len(r) else None

    def _tong_dong_con(key, j):
        """Σ 6 dòng con NGAY SAU dòng tổng `key` (nhãn con = nhãn tổng + hậu tố).

        Vì sao cần: kế toán có lúc bỏ trống Ô TỔNG mà vẫn điền các dòng con — lúc đó dòng tổng
        đọc ra 0 và cả dự án MẤT TRẮNG khỏi báo cáo ngày trong im lặng. Cộng dòng con là dựng lại
        đúng con số nghiệp vụ, KHÔNG phải suy đoán: 6 dòng con đúng bằng định nghĩa của dòng tổng
        ("Doanh thu thực hiện dự án" + "bán dầu thầu phụ" + "bán vật tư thầu phụ" + "chênh lệch
        kiểm kê dầu" + "điều chỉnh công nợ kê khai" + "khác").

        CHỈ cộng các dòng NGAY SAU dòng tổng và có nhãn BẮT ĐẦU bằng nhãn tổng, dừng ở dòng đầu
        tiên không khớp — để không trượt sang khối kế tiếp (III nằm ngay sau II) và không cộng
        đôi với chính dòng tổng.
        """
        r0 = anchored.get(key)
        if r0 is None:
            return None
        try:
            i0 = rows.index(r0)
        except ValueError:
            return None
        pref = _DUAN_ANCHOR[key][0]
        tong, thay = 0.0, False
        for r in rows[i0 + 1:]:
            if not r or ten_j >= len(r):
                break
            n = _nd(r[ten_j])
            if not n.startswith(pref) or n == pref:
                break
            v = _num(r[j]) if j < len(r) else None
            if v:
                tong += v
                thay = True
        return tong if thay else None

    def _duan_dt(j):
        """Doanh thu thuần của một dự án, có hai đường lùi.

        Thứ tự: (1) ô tổng III; (2) Σ 6 dòng con của III; (3) (I hoặc Σ con của I) − (II hoặc Σ
        con của II). Đường (3) đặt CUỐI vì nó suy từ khối GỘP — chỉ dùng khi cả III và các con
        của III đều trống, tránh lấy số gộp làm số thuần khi nguồn chỉ lỡ bỏ trống ô tổng.
        """
        v = val("dt_thuan", j)
        if v:
            return v
        v = _tong_dong_con("dt_thuan", j)
        if v:
            return v
        g = val("dt_gross", j) or _tong_dong_con("dt_gross", j)
        if not g:
            return None
        gt = val("giam_tru", j) or _tong_dong_con("giam_tru", j) or 0.0
        return g - gt

    facts = []
    for cc, j in cols:
        dt = _duan_dt(j)
        if dt:
            facts.append((cc, RT_HQKD, MA_DT, MA_DT, dt))
            facts.append((cc, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
            facts.append((cc, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
        gv = val("gia_von", j)
        if gv:
            facts.append((cc, RT_PNLT, "Giá vốn hàng bán", "Giá vốn hàng bán", gv))
        for kw, ten in _DUAN_GV_CT:
            r_gv = gv_ct.get(kw)
            v = _num(r_gv[j]) if r_gv is not None and j < len(r_gv) else None
            if v:
                facts.append((cc, RT_DUAN_GV, ten, ten, v, "Giá vốn"))
        ln_gop = val("ln_gop", j)
        if ln_gop:
            facts.append((cc, RT_PNLT, "Lợi nhuận gộp", "Lợi nhuận gộp", ln_gop))
        tong_cp = 0.0
        for key, nhom, ten in _DUAN_CP:
            v = val(key, j)
            if not v:
                continue
            tong_cp += v
            facts.append((cc, RT_CHIPHI, nhom, ten, v))
        if tong_cp:
            facts.append((cc, RT_HQKD, MA_CP, MA_CP, tong_cp))
        lntt = val("lntt", j)
        if lntt:
            facts.append((cc, RT_HQKD, MA_LNTT, MA_LNTT, lntt))
            facts.append((cc, RT_PNLT, "Lợi nhuận trước thuế", "Lợi nhuận trước thuế", lntt))
        lnst = val("lnst", j)
        if lnst:
            facts.append((cc, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lnst))
    return facts


_FACTS_FN = {"srvf": _srvf_facts, "kqkd": _kqkd_facts, "antaxi": _antaxi_facts, "tcode": _tcode_facts,
             "duan": _duan_facts, "ht": _ht_facts, "xdv": _xdv_facts}


# ---------------------------------------------------------------------------------------------
# Layout "ho_kqkd" — HO (HO/baocaohqkdngay), sheet "D1".."D31" luỹ kế GIỐNG HỆT "tcode" (GA/
# TRAMSAC) nhưng KHÔNG dùng mã T-series: cột A ("Mã số") của 3 DÒNG TỔNG (Doanh thu/Chi phí/Lợi
# nhuận) TRỐNG, chỉ có mã ở các dòng CHI TIẾT (511_TS/5118/515.01... — mã tài khoản kế toán, không
# phải mã báo cáo) -> PHẢI neo theo NHÃN cột B ("Chỉ tiêu"), không theo mã như "tcode".
# Verify file thật (kiểm 2026-08-05, kỳ 2026-08 D1/D2): header dòng 5: Mã số | Chỉ tiêu | D01 |
# %DT | T07 | %DT | T08 | … | T12 | %DT — cột GIÁ TRỊ luôn ngay SAU "Chỉ tiêu" (khớp header dạng
# d\d{2} zero-pad, dò bằng CHÍNH cơ chế val_j của "tcode"). Cột T07-T12 là TEMPLATE RÁC (mọi kỳ,
# mọi sheet đều =0/#DIV0!/#REF!, không đổi theo ngày thật) -> CHỈ lấy cột D01 (ngay sau Chỉ tiêu).
# 3 dòng tổng (không mã): "Tổng Doanh thu" (dòng 6) / "Tổng chi phí" (dòng 14) / "Tổng lợi nhuận"
# (dòng 26) — user tự ghi rõ ô C6/C14/C26 trong Mapping, nhưng neo THEO NHÃN thay vì địa chỉ ô
# cứng (số dòng có thể lệch giữa các tháng khi kế toán chèn/xoá dòng chi tiết).
# LNTT/LNST: verify DB 2026-01..07 — PNLT của HO CHỈ có đúng 1 dim1 'Lợi nhuận sau thuế', giá trị
# BẰNG TUYỆT ĐỐI HQKD mã 1112 ở CẢ 7 kỳ (vd T07: -4,5950 = -4,5950) -> 'Tổng lợi nhuận' ĐÃ LÀ LNST
# (giống Trạm sạc, KHÁC GA) -> dùng profit_pnlt=('Lợi nhuận sau thuế',) như Trạm sạc.
# KHÔNG 'Giá vốn hàng bán'/'Lợi nhuận gộp': HO là khối hỗ trợ tập đoàn (chi phí quản lý/lãi vay),
# monthly KHÔNG có dim1 'Giá vốn hàng bán' cho HO (đơn vị không có COGS) — daily cũng không suy.
# Cơ cấu chi phí (CHIPHI_D): monthly HIỆN CHƯA CÓ breakdown nào cho HO (0 dòng CHIPHI) nhưng
# Mapping yêu cầu rõ (#10: "lấy từng loại chi phí trong file BC (C15-C25)") và file NGÀY CÓ sẵn 11
# dòng chi tiết (511_TS/5118/… ở phần DT không tính, CHỈ 11 dòng CHI PHÍ nằm giữa 'Tổng chi phí' và
# 'Tổng lợi nhuận') -> lấy NHÃN GỐC (cột B) làm dim1 THẲNG (không có nhóm chuẩn nào để map vào vì
# monthly chưa từng phân loại) — daily "đi trước" monthly ở khoản này, chấp nhận, không tự đặt tên
# nhóm mới có thể sai. Quét theo VÙNG (dòng SAU anchor 'Tổng chi phí', TRƯỚC anchor 'Tổng lợi
# nhuận'), KHÔNG hardcode "dòng 15-25" vì số dòng chi tiết có thể đổi giữa các tháng.
# ⚠ SỬA 2026-08-06 — "Doanh thu" KHÔNG lấy nguyên văn dòng "Tổng Doanh thu": dòng đó trong FILE
# THẬT cộng cả 515.xx (DT tài chính) + 7111 (thu nhập khác). Quy ước ĐÃ CHỐT với user 2026-07-30 cho
# bản THÁNG (agent_cli._derive_kqkd_ho::code511_sum, QA #10 T2/2026: dòng 6 = 32.833.798 ≠ Σ511 =
# 30.832.363 là số kế toán mong đợi) là **Doanh thu = Σ MỌI dòng mã 511*** (511_TS · 5118 · 5119 ·
# 511_LN). Bản NGÀY trước đây neo nhầm vào dòng tổng -> sẽ thổi phồng doanh thu HO đúng bằng
# 515+7111 ngay khi file có số. Anchor 'dt' giữ lại CHỈ để fallback khi sheet không có cột mã.
_HO_ANCHOR = {"dt": "tong doanh thu", "tong_cp": "tong chi phi", "lntt": "tong loi nhuan"}


def _ho_facts(rows):
    """rows -> [(None, report_type, dim1, dim3, value_VND)] cho 1 ngày. [] nếu sai layout."""
    hdr_i = next((i for i, r in enumerate(rows[:10]) if any(_nd(c) == "chi tieu" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    ten_j = next(j for j, c in enumerate(hdr) if _nd(c) == "chi tieu")
    val_j = next((j for j, c in enumerate(hdr) if j > ten_j and re.fullmatch(r"d\d{2}", _nd(c))),
                 ten_j + 1)

    anchor_i = {}
    for i, r in enumerate(rows[hdr_i + 1:], hdr_i + 1):
        if not r or ten_j >= len(r):
            continue
        n = _nd(r[ten_j])
        for key, pref in _HO_ANCHOR.items():
            if key not in anchor_i and n == pref:
                anchor_i[key] = i
    if "dt" not in anchor_i or "tong_cp" not in anchor_i or "lntt" not in anchor_i:
        return []

    def val(key):
        i = anchor_i.get(key)
        r = rows[i] if i is not None else None
        return _num(r[val_j]) if r is not None and val_j < len(r) else None

    # Doanh thu = Σ dòng mã 511* (xem chú thích ở _HO_ANCHOR). Cột mã = header 'Mã số', fallback cột
    # ngay TRƯỚC 'Chỉ tiêu'. Excel lưu mã dạng số ('5118' -> 5118.0) nên phải bỏ đuôi '.0'.
    ma_j = next((j for j, c in enumerate(hdr) if _nd(c) == "ma so"), ten_j - 1)
    dt_511, co_511 = 0.0, False
    if ma_j >= 0:
        for r in rows[hdr_i + 1:]:
            if not r or ma_j >= len(r) or r[ma_j] in (None, ""):
                continue
            code = str(r[ma_j]).strip()
            if code.endswith(".0"):
                code = code[:-2]
            v = _num(r[val_j]) if val_j < len(r) else None
            if code.startswith("511") and v is not None:
                dt_511 += v
                co_511 = True

    facts = []
    # Không dòng 511* nào có số -> rơi về dòng "Tổng Doanh thu" (sheet đổi layout, mất cột mã).
    dt = dt_511 if co_511 else val("dt")
    if dt:
        facts.append((None, RT_HQKD, MA_DT, MA_DT, dt))
        facts.append((None, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
        facts.append((None, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
        # LN gộp = DT thuần (HO không có giá vốn) — Y HỆT bản THÁNG, chốt Mapping 2026-07-18.
        # Bản ngày trước đây THIẾU dòng này -> thẻ "Lợi nhuận gộp" của HO luôn trống ở chế độ Ngày.
        facts.append((None, RT_PNLT, "Lợi nhuận gộp", "Lợi nhuận gộp", dt))
    tong_cp = val("tong_cp")
    if tong_cp:
        facts.append((None, RT_HQKD, MA_CP, MA_CP, tong_cp))
    lntt = val("lntt")
    if lntt:
        facts.append((None, RT_HQKD, MA_LNTT, MA_LNTT, lntt))
        facts.append((None, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lntt))
    # Cơ cấu chi phí: mọi dòng có nhãn NẰM GIỮA anchor 'tong_cp' và 'lntt' -> 1 dim1 riêng theo
    # nhãn gốc (xem docstring khối trên).
    for i in range(anchor_i["tong_cp"] + 1, anchor_i["lntt"]):
        r = rows[i]
        if not r or ten_j >= len(r):
            continue
        ten = str(r[ten_j]).strip() if r[ten_j] not in (None, "") else None
        v = _num(r[val_j]) if val_j < len(r) else None
        if ten and v:
            facts.append((None, RT_CHIPHI, ten, ten, v))
    return facts


_FACTS_FN["ho_kqkd"] = _ho_facts


# ---------------------------------------------------------------------------------------------
# Layout "anks" — An KS (ANKHACHSAN), KHÁC HẲN 3 layout trên: chỉ 1 SHEET DUY NHẤT, mỗi NGÀY là
# 1 CỘT (không phải mỗi ngày 1 sheet). Header dòng 3: TT | Nội dung chi phí | ĐVT | Tổng cộng |
# 1 | 2 | … | 31 (số ngày nằm NGAY Ở HEADER, không cần suy). Nhãn chỉ tiêu ở cột "Nội dung chi
# phí", neo theo TIỀN TỐ đã chuẩn hoá (mã La Mã I/II/III ở cột riêng bên trái, không dính nhãn).
# CỘT "Tổng cộng" = Σ các cột ngày trong tháng (verify ngày 1-4/08: 2.975.000+1.600.000+1.615.000
# +1.400.000 = 7.590.000 = đúng ô Tổng cộng dòng I) -> BỎ cột này, chỉ lấy cột NGÀY để không đếm
# đôi khi gộp nhiều ngày (giống lý do bỏ cột "Tổng cộng" ở layout kqkd).
# KHÔNG cost center: An KS chỉ 1 cơ sở (Garden Sơn Tây) — verify raw_rows bản THÁNG: cột
# cost_center của khối An KS luôn NULL ở mọi report_type/mọi kỳ.
# dim1 PNLT/CHIPHI lấy Y HỆT bản THÁNG của An KS (rà DB 2026-06):
#   CHIPHI: 'Giá vốn hàng bán' / 'Chi phí chung' / 'Chi phí lương + CP khác cho CNV' (3 nhóm,
#     KHÁC An Taxi/XVP) — đúng 3 dòng con II.1/II.2/II.3 của file.
#   PNLT: ghi CẢ HAI tên gross ('Doanh thu HH, DV' VÀ 'Doanh thu bán hàng và cung cấp dịch vụ',
#     cùng giá trị — như An Taxi, xem doanhthu-cautructructure-gross-giamtru) + 'Giá vốn hàng
#     bán' + CHỈ 'Lợi nhuận sau thuế' (An KS KHÔNG có dòng thuế TNDN -> LNTT=LNST, bản THÁNG
#     không ghi dim1 'Lợi nhuận trước thuế' riêng, xem An.xlsx spec — daily giữ đúng convention,
#     KHÔNG tự thêm dòng monthly không có). 'Lợi nhuận gộp' KHÔNG ghi tường minh — bản THÁNG lưu
#     nó nhưng do BACKEND tự tính (revenue.py::pnl.py fallback DT-giá vốn theo khối khi thiếu
#     dòng riêng), daily để trống cho backend tự suy, tránh trùng 2 nguồn.
_ANKS_ANCHOR = {
    "dt": "tong doanh thu",              # I
    "tong_cp": "tong chi phi",           # II (nhãn gốc có dấu ':' cuối -> startswith vẫn khớp)
    "gia_von": "chi phi gia von",        # II.1
    "chi_chung": "chi phi chung",        # II.2
    "luong": "chi phi luong",            # II.3 "CHI PHÍ LƯƠNG + CP KHÁC CHO CNV"
    "lntt": "loi nhuan (i-ii)",          # III — An KS: LNTT = LNST (không có dòng thuế TNDN)
}


def _anks_all_days(rows, period):
    """1 sheet, mỗi cột 1 ngày -> [(ngay, facts), ...] cho TRỌN THÁNG trong 1 lần gọi (khác hẳn
    3 layout trên vốn mỗi ngày gọi facts_fn riêng — xem docstring khối trên)."""
    hdr_i = next((i for i, r in enumerate(rows[:10])
                  if any(_nd(c) == "tong cong" for c in r if c is not None)), None)
    if hdr_i is None:
        return []
    hdr = rows[hdr_i]
    tong_j = next(j for j, c in enumerate(hdr) if _nd(c) == "tong cong")
    y, mm = int(period[:4]), int(period[5:7])
    day_cols = [(j, int(c)) for j, c in enumerate(hdr) if j > tong_j and isinstance(c, int)
                and not isinstance(c, bool) and 1 <= c <= 31]

    label_j = next((j for j, c in enumerate(hdr) if _nd(c).startswith("noi dung")), 2)
    anchored = {}
    for r in rows[hdr_i + 1:]:
        if not r or label_j >= len(r):
            continue
        n = _nd(r[label_j])
        for key, pref in _ANKS_ANCHOR.items():
            if key not in anchored and n.startswith(pref):
                anchored[key] = r

    def val(key, j):
        r = anchored.get(key)
        return _num(r[j]) if r is not None and j < len(r) else None

    per_day = []
    for j, d in day_cols:
        facts = []
        dt = val("dt", j)
        if dt:
            facts.append((None, RT_HQKD, MA_DT, MA_DT, dt))
            facts.append((None, RT_DTHU, "Doanh thu thuần", "Doanh thu thuần", dt))
            facts.append((None, RT_PNLT, "Doanh thu HH, DV", "Doanh thu HH, DV", dt))
            facts.append((None, RT_PNLT, "Doanh thu bán hàng và cung cấp dịch vụ",
                          "Doanh thu bán hàng và cung cấp dịch vụ", dt))
        tong_cp = val("tong_cp", j)
        if tong_cp:
            facts.append((None, RT_HQKD, MA_CP, MA_CP, tong_cp))
        for key, dim1 in (("gia_von", "Giá vốn hàng bán"), ("chi_chung", "Chi phí chung"),
                          ("luong", "Chi phí lương + CP khác cho CNV")):
            v = val(key, j)
            if v:
                facts.append((None, RT_CHIPHI, dim1, dim1, v))
        gia_von = val("gia_von", j)
        if gia_von:
            facts.append((None, RT_PNLT, "Giá vốn hàng bán", "Giá vốn hàng bán", gia_von))
        lntt = val("lntt", j)
        if lntt:
            facts.append((None, RT_HQKD, MA_LNTT, MA_LNTT, lntt))
            facts.append((None, RT_PNLT, "Lợi nhuận sau thuế", "Lợi nhuận sau thuế", lntt))
        if facts:
            per_day.append((f"{y:04d}-{mm:02d}-{d:02d}", facts))
    return per_day


def _ban_sinh_doi(path):
    """Có file VẬT LÝ KHÁC cùng basename ở thư mục report_type khác của cùng công ty không?

    `raw_rows.source_file` chỉ là `<FOLDER_CÔNG_TY>::<basename>` — KHÔNG kèm report_type. An Taxi
    có `ANTAXI/baocaohqkdngay/B.7.AAG.TCKT.M.202608.Baocaotaichinhrieng.xlsx` (nguồn NGÀY, 281 KB)
    và `ANTAXI/baocaotaichinhrieng/` cùng TÊN Y HỆT (BCTC THÁNG, 21,6 MB) -> cùng một khoá. Xoá
    trần theo khoá đó là mỗi lượt cron ngày quét sạch dòng BCTC tháng của đơn vị (đo 16/09 và lại
    18/09/2026: An Taxi mất trắng số tháng 8 ở toàn cụm tài chính).
    """
    try:
        thumuc_rt = os.path.dirname(os.path.abspath(path))
        thumuc_cty = os.path.dirname(thumuc_rt)
        ten = os.path.basename(path)
        return any(os.path.isfile(os.path.join(thumuc_cty, d, ten))
                   for d in os.listdir(thumuc_cty)
                   if os.path.join(thumuc_cty, d) != thumuc_rt)
    except OSError:
        return False        # không đọc được đĩa -> giữ nguyên hành vi cũ (xoá trần)


# ---------------------------------------------------------------------------------------------
def derive(path, write=False):
    folder = _source_id(path).split("::", 1)[0]
    unit = _UNITS.get(folder)
    period = _period_of(os.path.basename(path), _in_day_dir(path))
    if not unit or not period:
        return {"ok": False, "skip": True}

    # FILE HỎNG / RỖNG -> câu lý do rõ, KHÔNG để traceback thoát ra (thêm 06/09/2026). Ca thật:
    # `TRAMSAC/baocaohqkdngay/B.3.TC.TCKT.D.20260829.Baocaotaichinhrieng.xlsx` nặng ĐÚNG 7 byte
    # (không phải zip) -> openpyxl ném BadZipFile -> `cron_hqkdngay_daily.autofill()` chỉ ghi được
    # `rc=1` trần vào log, không ai biết vì sao. Đây là lỗi PHÍA NGUỒN (file truyền dở), không
    # phải lỗi layout, nên phải nói khác hẳn để người đọc log không đi sửa nhầm chỗ.
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        kb = os.path.getsize(path) / 1024 if os.path.exists(path) else 0
        return {"ok": False, "period": period, "layout": unit["layout"],
                "file_hong": True,
                "error": f"FILE KHÔNG MỞ ĐƯỢC ({kb:.1f} KB): {type(e).__name__}. Nguồn gửi dở/hỏng, "
                         f"cần kéo lại — không phải lỗi layout"}
    sheets = None            # None = layout "anks" (1 sheet, không dò theo ngày)
    snap_mode = False        # True = bộ snapshot luỹ kế (xem `_tcode_snap_per_day`)
    bcqt_mode = False        # True = họ file "BCQT PT." của An Taxi (xem `_bcqt_per_day`)
    # LAYOUT THEO FILE, KHÔNG THEO ĐƠN VỊ: An Taxi có hai họ file nằm LẪN trong cùng thư mục nên
    # `_UNITS[folder]["layout"]` không đủ để quyết. Nhận theo NỘI DUNG (sheet nào có mặt), cùng
    # nguyên tắc `snapshot_sheet` của Trạm sạc. Đơn vị không khai `layout_phu` thì không đổi gì.
    layout = unit["layout"]
    phu = unit.get("layout_phu")
    if phu and phu[0] in (wb.sheetnames or []):
        layout = phu[1]
    try:
        if layout == "antaxi_bcqt":
            per_day = []          # dựng sau khi đóng workbook — `_bcqt_per_day` tự mở cả bộ file
            bcqt_mode = True
        elif layout == "anks":
            rows = [list(r) for r in wb[wb.sheetnames[0]].iter_rows(values_only=True)]
            per_day = _anks_all_days(rows, period)
        else:
            facts_fn = _FACTS_FN[layout]
            if layout == "tcode":
                # 2 đơn vị chung layout "tcode" (GA/TRAMSAC) nhưng quy ước PNLT LNTT/LNST khác
                # nhau — xem docstring `_tcode_facts` param `profit_pnlt`.
                pp = unit.get("profit_pnlt", ("Lợi nhuận trước thuế",))
                ps = unit.get("pnlt_skip", ())
                facts_fn = lambda rows, _pp=pp, _ps=ps: _tcode_facts(rows, _pp, _ps)  # noqa: E731
            if layout == "srvf":
                sheets = _srvf_day_sheets(wb, period)
            elif layout in ("tcode", "ho_kqkd"):   # cùng kiểu sheet "D1".."D31" luỹ kế
                sheets = _tcode_day_sheets(wb, period)
                # KHÔNG CÓ sheet ngày + đơn vị có khai `snapshot_sheet` -> bộ file mẫu-tháng-chụp-
                # theo-ngày (Trạm sạc từ kỳ 2026-09). Nhận diện bằng NỘI DUNG chứ không bằng mốc
                # ngày: file cũ/mới nằm lẫn trong cùng thư mục, và chính docstring `_UNITS` đã ghi
                # vì sao mốc `bo_tu_ngay` là hướng sai cho đơn vị này.
                if not sheets and unit.get("snapshot_sheet") in (wb.sheetnames or []):
                    snap_mode = True
            elif layout in ("ht", "xdv"):          # "{dd}.{mm}" hoặc "{dd}" trần
                sheets = _ht_day_sheets(wb, period)
            else:
                sheets = _kqkd_day_sheets(wb, period)
            per_day = []
            for sheet, ngay in sheets:
                rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
                facts = facts_fn(rows)
                if facts:
                    per_day.append((ngay, facts))
    finally:
        wb.close()

    # SNAPSHOT LUỸ KẾ: phải mở THÊM các file khác cùng kỳ để trừ ra số ngày -> làm sau khi đã
    # đóng workbook của chính file này, tránh giữ 2 handle cùng lúc trên cùng một file.
    snap_diag = {}
    if snap_mode:
        per_day, snap_diag = _tcode_snap_per_day(path, period, unit)
    if bcqt_mode:
        per_day, snap_diag = _bcqt_per_day(path, period, unit)

    # SỐ DƯ THEO NGÀY CHO ĐƠN VỊ KHÔNG Ở CHẾ ĐỘ SNAPSHOT (3 khối Xanh, 18/09/2026) — P&L và số dư
    # nằm ở HAI HỌ FILE tách rời nên phải quét thư mục thêm một lượt, xem `_sodu_quet_thu_muc`.
    #
    # CHỈ FILE P&L ĐƯỢC GỘP (`_snap_day_of` trả None = tên chỉ có 6 số sau `.D.`). Bỏ điều kiện
    # này thì mỗi file ảnh chụp cũng tự ghi một bộ số dư dưới `source_file` của RIÊNG nó: cùng một
    # ngày nằm trong DB dưới nhiều khoá, và `DELETE ... WHERE source_file=%s` của lượt sau không
    # dọn nổi dòng của lượt trước — số dư nhân lên theo số file, âm thầm.
    #
    # Gộp SAU `bo_tu_ngay`/`bo_ngay_tuong_lai`? KHÔNG — gộp ở ĐÂY để hai mốc đó cắt được cả dòng
    # số dư. Mốc là thứ tách hai NGUỒN của cùng một đơn vị, số dư không được miễn trừ.
    #
    # GATE LÀ CỜ RIÊNG `sodu_ho_file_rieng`, KHÔNG PHẢI `sheet_sodu`. An Taxi cũng khai
    # `sheet_sodu` (cho `_bcqt_per_day` dùng) và file họ `.M.` của nó KHÔNG vào `bcqt_mode` -> gate
    # theo `sheet_sodu` là lượt nạp file `.M.` đi quét thư mục và gom nhầm 16 ngày số dư của họ
    # `.D.`, ghi đè dưới `source_file` của file `.M.` trong khi `bcqt_mode` đã ghi chúng dưới khoá
    # THEO KỲ. Đo thật 18/09/2026: đúng 16 ngày bị gom. Hôm nay chưa hỏng chỉ vì `bo_tu_ngay`
    # 01/09 cắt sạch chúng ngay sau đó — tức an toàn nhờ một con số không liên quan.
    sodu_diag = {}
    if (unit.get("sodu_ho_file_rieng") and not (snap_mode or bcqt_mode)
            and not _snap_day_of(os.path.basename(path))):
        gop_sd, bo_qua_sd, luy_ke_sd = _sodu_quet_thu_muc(path, period, unit)
        if gop_sd:
            theo_ngay = {n: list(f) for n, f in per_day}
            for ngay, facts in gop_sd.items():
                theo_ngay.setdefault(ngay, []).extend(facts)
            per_day = sorted(theo_ngay.items())
        sodu_diag = {"so_du_theo_ngay": {
            "so_ngay": len(gop_sd), "ngay": sorted(gop_sd),
            **({"ngay_tu_ban_luy_ke": luy_ke_sd} if luy_ke_sd else {}),
            **({"bo_qua_file": bo_qua_sd} if bo_qua_sd else {})}}

    # P&L VÉT TỪ HỌ FILE ẢNH CHỤP (19/09/2026) — cùng cổng và cùng lý do với cụm số dư ngay trên:
    # chỉ chạy ở lượt nạp file THÁNG, để mọi dòng nằm dưới một `source_file` duy nhất và lệnh
    # `DELETE ... WHERE source_file=%s` dọn được trọn vẹn. Xem `_pl_quet_thu_muc`.
    pl_diag = {}
    if (unit.get("pl_ho_file_rieng") and not (snap_mode or bcqt_mode)
            and not _snap_day_of(os.path.basename(path))):
        da_co = {n for n, f in per_day if any(x[1] in _RT_DONG_CHAY for x in f)}
        gop_pl, bo_qua_pl = _pl_quet_thu_muc(path, period, unit, da_co)
        if gop_pl:
            theo_ngay = {n: list(f) for n, f in per_day}
            for ngay, facts in gop_pl.items():
                theo_ngay.setdefault(ngay, []).extend(facts)
            per_day = sorted(theo_ngay.items())
        pl_diag = {"pl_tu_anh_chup": {
            "so_ngay": len(gop_pl), "ngay": sorted(gop_pl),
            **({"bo_qua_file": bo_qua_pl} if bo_qua_pl else {})}}

    # CUTOVER SANG NGUỒN TỰ ĐỘNG (xem `bo_tu_ngay` trong `_UNITS`): cắt TRƯỚC nhánh báo lỗi bên
    # dưới và TRƯỚC lệnh ghi, để file vẫn đi trọn đường xuống `DELETE ... WHERE source_file=%s`.
    # Bỏ sớm bằng `return` là những ngày đã nạp trước mốc còn nằm lại trong DB và cộng đôi với
    # nguồn mới — đúng cái bẫy mốc cutover sinh ra để chặn.
    #
    # CHỈ ÁP CHO LAYOUT GỐC của đơn vị: mốc là thứ tách hai NGUỒN của cùng một đơn vị (An Taxi —
    # họ `.M.` nhường họ `.D.` từ 01/09/2026), nên áp cả cho nguồn MỚI là nguồn mới tự chặn chính
    # mình và đơn vị mất sạch số từ mốc trở đi. Đơn vị không khai `layout_phu` thì điều kiện này
    # luôn đúng -> hành vi cũ của Showroom/XDV không đổi.
    moc = unit.get("bo_tu_ngay") if layout == unit["layout"] else None
    bo_cutover = None
    if moc:
        cat = sorted(n for n, _ in per_day if n >= moc)
        if cat:
            per_day = [(n, f) for n, f in per_day if n < moc]
            bo_cutover = {"tu_ngay": moc, "so_ngay_bo": len(cat), "dau": cat[0], "cuoi": cat[-1]}

    # NGÀY CHƯA TỚI (chốt 15/09/2026): mọi `_*_day_sheets` ở trên suy ngày từ TÊN SHEET, mà nguồn
    # tạo sẵn sheet của ngày chưa tới — HTX Xanh Vĩnh Phúc sáng 15/09 đã có sheet "16", ruột là bản
    # copy của template nên doanh thu bằng y hệt mấy ngày trước (6.139.871 đ lặp ở 01/02/04/06/07/
    # 12→16). Không có dấu hiệu nào trong RUỘT sheet để phân biệt "bản copy chờ điền" với số thật,
    # nên chặn bằng mốc ngày. Cắt ở ĐÂY, cùng chỗ với `bo_tu_ngay` và trước lệnh ghi, để file vẫn
    # đi trọn đường xuống `DELETE ... WHERE source_file=%s` — bỏ sớm bằng `return` thì dòng ngày
    # tương lai nạp hôm trước còn nằm lại trong DB.
    # Giờ VN cố định +07: crontab máy này chạy UTC, `date.today()` từ 00:00-07:00 giờ VN còn đứng ở
    # hôm qua và sẽ cắt nhầm đúng ngày mới nhất.
    hom_nay = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))).date().isoformat()
    bo_tuong_lai = None
    tl = sorted(n for n, _ in per_day if n > hom_nay)
    if tl:
        per_day = [(n, f) for n, f in per_day if n <= hom_nay]
        bo_tuong_lai = {"hom_nay": hom_nay, "so_ngay_bo": len(tl), "dau": tl[0], "cuoi": tl[-1]}

    if not per_day and not bo_cutover and not bo_tuong_lai:
        # ĐƠN VỊ DÙNG CHÍNH FILE NÀY LÀM NGUỒN SỐ DƯ -> KHÔNG PHẢI LỖI (18/09/2026).
        #
        # Các đơn vị có hai họ file trong cùng thư mục (3 khối Xanh, An Khách sạn): P&L ở
        # `...D.<YYYYMM>.`, ảnh chụp số dư ở `...D.<YYYYMMDD>.`. Lượt nạp file P&L đã quét cả bộ ảnh
        # chụp và ghi số dư (xem `_sodu_quet_thu_muc`), nên tới lượt CHÍNH file ảnh chụp thì đúng là
        # nó không còn gì để ghi — đó là thiết kế, không phải hỏng.
        #
        # Báo lỗi ở đây làm `cron_hqkdngay_daily.autofill()` trả not-ok -> bảng giám sát ghim đơn vị
        # ở `loi_nap` kèm câu "cần kế toán kiểm tra lại file", tức BẢO KẾ TOÁN SỬA MỘT FILE ĐANG
        # ĐÚNG. Đo 18/09/2026: cả 3 khối Xanh + An KS đỏ vì đúng lý do này.
        #
        # ĐẶT TRƯỚC MỌI NHÁNH LỖI và KHÔNG phụ thuộc layout: An KS đi layout "anks" (`sheets` là
        # None) nên không rơi vào nhánh `la_mau_thang` của các khối Xanh, phải bắt ở tầng chung.
        #
        # Vẫn PHẢI CHỨNG MINH file dùng được, không tin suông cấu hình: đọc lại kỳ trong sheet số dư
        # của chính nó. Không ra ngày hợp lệ thì rơi tiếp xuống các nhánh lỗi — file hỏng vẫn đỏ.
        if unit.get("sodu_ho_file_rieng") and _snap_day_of(os.path.basename(path)):
            wb2 = None
            try:
                wb2 = openpyxl.load_workbook(path, data_only=True, read_only=True)
                ngay_sd, vi_sao_sd = _sodu_ngay_cua_wb(wb2, unit.get("sheet_sodu") or {},
                                                       os.path.basename(path),
                                                       bool(unit.get("ky_sodu_o_ngay")))
            except Exception as e:
                ngay_sd, vi_sao_sd = None, f"không mở được ({type(e).__name__})"
            finally:
                if wb2 is not None:
                    wb2.close()
            if ngay_sd:
                return {"ok": True, "file": os.path.basename(path), "period": period,
                        "cong_ty": unit["cong_ty"], "layout": layout, "days": 0,
                        "la_nguon_so_du": {"ngay": ngay_sd, "nap_boi": "file P&L cùng kỳ"},
                        "tong_theo_ngay": {}}
            # Không nhận được ngày -> nói ĐÚNG lý do của cụm số dư. Rơi xuống nhánh chung bên dưới
            # thì câu lỗi nói về sheet NGÀY ("không thấy sheet ngày nào khớp kỳ — layout anks"),
            # dẫn người đọc đi tìm sai chỗ: file này không bao giờ có sheet ngày, nó là ảnh chụp.
            return {"ok": False, "period": period, "layout": layout, "la_anh_chup_so_du": True,
                    "error": f"ảnh chụp số dư không dùng được: {vi_sao_sd}"}
        # PHÂN BIỆT 3 nguyên nhân — bản đầu gộp chung 1 câu "không đọc được sheet ngày nào" khiến
        # chẩn đoán đi nhầm hướng (2026-08-06, Trạm sạc/GA/HO kỳ 08: tưởng hỏng dò sheet, hoá ra
        # sheet đọc tốt nhưng file nguồn ghi 0 CỨNG ở mọi mã tổng — xem `sheets_ngay` trả kèm).
        #
        # NGUYÊN NHÂN THỨ BA, thêm 06/09/2026: FILE ĐỔI HẲN LAYOUT — bỏ dải sheet ngày, thay bằng
        # MỘT sheet luỹ kế của bản báo cáo tự động ("BCHQKD"/"BCĐPS"/"BCKQKD"...). Ca thật: Trạm
        # sạc từ bản `.D.20260829.` trở đi. Hai câu lý do cũ đều dẫn sai đường — "không thấy sheet
        # ngày nào khớp kỳ" đọc như hỏng phép dò sheet (tôi đã mất một vòng debug đúng vào đó), và
        # trên bảng giám sát nó ra "kế toán kiểm tra lại file" trong khi việc phải làm là hỏi bên
        # sinh file vì sao đổi khuôn và vì sao P&L toàn 0.
        #
        # `loi_nap` VẪN LÀ TRẠNG THÁI ĐÚNG (file có, số không có) — cố ý không đổi sang một trạng
        # thái "lành": không có nguồn nào khác ghi số cho những ngày này, tô xanh là che một luồng
        # dữ liệu đã đứt. Chỉ sửa CÂU LÝ DO cho trỏ đúng chỗ.
        if bcqt_mode:
            # `sheets` là None ở nhánh này (không dò theo tên sheet) nên hai câu dưới đều không
            # khớp; không có câu riêng thì rơi xuống "không thấy sheet ngày nào khớp kỳ" — đọc như
            # hỏng phép dò sheet trong khi sheet vẫn ở đó, chỉ là mọi file đều bị loại.
            bq = (snap_diag or {}).get("bo_qua_file") or []
            return {"ok": False, "period": period, "layout": layout, **(snap_diag or {}),
                    "error": f"sheet '{_BCQT_SHEET}' có nhưng không ra ngày nào cho kỳ {period}"
                             + (f" — {len(bq)} file bị loại: {bq[0]['vi_sao']}" if bq else
                                " — cột ngày còn trống, kế toán chưa xuất số")}
        if sheets is not None and not sheets:
            # `_nd` gộp khoảng trắng nhưng GIỮ dấu chấm/cách ("BCQT PT." -> "bcqt pt."), nên phải
            # bỏ nốt ký tự không phải chữ/số trước khi so.
            # HAI HỌ TÊN SHEET của bản tự động, cùng một bộ báo cáo nhưng đặt tên khác nhau —
            # nhận cả hai, nếu không thì câu lý do liệt kê thiếu và đọc như "file lạ hoàn toàn":
            #   · Trạm sạc  : BCHQKD · BCĐPS · BCKQKD · BCĐKT · BC LCTT   (có tiền tố "BC")
            #   · An Taxi   : KQKD · CDPS · CDKT · LCTT · BCQT PT. · DATA (không tiền tố)
            moi = [s for s in wb.sheetnames
                   if re.sub(r"[^a-z0-9]", "", _nd(s))
                   in ("bchqkd", "bcdps", "bckqkd", "bcdkt", "bcqtpt", "bclctt",
                       "kqkd", "cdps", "cdkt", "lctt")]
            if moi:
                # GIỮ CÂU DƯỚI 200 KÝ TỰ: `cron_hqkdngay_daily.autofill()` cắt `errs[0][:200]`
                # trước khi đưa vào `ly_do_ky_thuat` của bảng giám sát. Bản đầu dài ~330 ký tự nên
                # đúng phần HÀNH ĐỘNG ở cuối bị cắt mất, chỉ còn "Cần chốt: " lửng. Vì thế nêu
                # nguyên nhân + việc phải làm TRƯỚC, và chỉ liệt kê 2 sheet đầu.
                #
                # CÂU NÀY ĐÃ SỬA LẠI 06/09/2026 sau khi user chỉ ra: bản đầu viết "FILE ĐỔI
                # LAYOUT ... chỉ còn sheet luỹ kế bản tự động", đọc như một khuôn LẠ cần viết
                # parser mới — và tôi đã đi đúng vào hướng đó rồi phải gỡ. Sự thật khác: bộ sheet
                # này TRÙNG TUYỆT ĐỐI mẫu báo cáo THÁNG trong `baocaotaichinhrieng/` (đã nạp
                # nhiều kỳ), tức đây là mẫu THÁNG chụp theo ngày và số là LUỸ KẾ từ đầu tháng.
                # Nói đúng bản chất thì người đọc log không mất một vòng như tôi.
                return {"ok": False, "period": period, "layout": layout,
                        "doi_layout": moi, "la_mau_thang": True,
                        "error": f"KHÔNG PHẢI BÁO CÁO NGÀY kỳ {period}: file là mẫu báo cáo THÁNG "
                                 f"chụp theo ngày ({', '.join(moi[:2])}...), số LUỸ KẾ từ đầu "
                                 f"tháng — nạp vào đường ngày là phóng đại. Số ngày nằm ở sheet "
                                 f"sổ nhật ký"}
        if sheets is not None and sheets:
            names = ", ".join(s for s, _ in sheets)
            return {"ok": False, "period": period, "layout": layout,
                    "sheets_ngay": [s for s, _ in sheets],
                    "error": f"đọc được {len(sheets)} sheet ngày ({names}) nhưng không sheet nào ra "
                             f"chỉ tiêu có số — kiểm tra file nguồn còn để trống/ghi 0 ở các mã tổng "
                             f"(layout {layout}, kỳ {period})"}
        return {"ok": False, "error": f"không thấy sheet ngày nào khớp kỳ {period} "
                                      f"(layout {layout})"}

    out = {"ok": True, "file": os.path.basename(path), "period": period, "cong_ty": unit["cong_ty"],
           "layout": layout, "days": len(per_day),
           **({"bo_qua_cutover": bo_cutover} if bo_cutover else {}),
           **({"bo_ngay_tuong_lai": bo_tuong_lai} if bo_tuong_lai else {}),
           **snap_diag, **sodu_diag, **pl_diag,
           "tong_theo_ngay": {
               # `x[:5]` chứ không giải nén cứng 5 phần tử: fact của layout "srvf" có thêm
               # dim2 (kênh bán) ở vị trí thứ 6.
               ngay: {"doanh_thu_ty": round(sum(x[4] for x in f
                                                if x[1] == RT_HQKD and x[2] == MA_DT) * 1e-9, 9),
                      "chi_phi_ty": round(sum(x[4] for x in f
                                              if x[1] == RT_HQKD and x[2] == MA_CP) * 1e-9, 9),
                      "lntt_ty": round(sum(x[4] for x in f
                                           if x[1] == RT_HQKD and x[2] == MA_LNTT) * 1e-9, 9),
                      "ban_xe_ty": round(sum(x[4] for x in f if x[1] == RT_KDVH) * 1e-9, 9)}
               for ngay, f in per_day}}
    if not write:
        return out

    source_file = _source_id(path)
    # SNAPSHOT: khoá idempotent phải là KỲ, không phải FILE. Mỗi lượt nạp dựng lại TOÀN BỘ ngày
    # của kỳ từ cả bộ snapshot, nên nếu vẫn khoá theo tên file thì file 09/09 chỉ xoá được dòng
    # do chính nó ghi, còn dòng file 08/09 ghi hôm trước nằm lại -> mỗi ngày bị đếm 2 lần.
    # `bcqt_mode` y hệt, xem `_bcqt_per_day` — 12 file cùng kỳ đều khai đủ các ngày đã qua.
    if snap_mode or bcqt_mode:
        source_file = period_source_key(folder, period)
    conn = psycopg.connect(DB_URL)
    try:
        cur = conn.cursor()
        # KỲ: báo cáo ngày là SỐ THỰC TẾ của một kỳ ĐÃ TỚI -> được khai sinh kỳ nếu chưa có.
        # Trước 03/09/2026 chỗ này chỉ SELECT rồi bỏ cả file, nên đầu tháng nào số ngày cũng nằm
        # chờ tới khi báo cáo THÁNG về mới có kỳ để treo vào (kỳ 2026-08 sinh ngày 04/08; kỳ
        # 2026-09 làm 5 đơn vị mất số ngày 01-03/09 dù file đã về đĩa từ 16:36 ngày 03/09).
        # Vì sao được phép tạo, và vì sao KẾ HOẠCH thì không: xem servers/common/dataset_ky.py.
        dataset_id, _tt = _DSK.lay_hoac_tao_ky(cur, period, nguon=os.path.basename(path))
        if not dataset_id:
            out["ok"] = False
            out["error"] = _DSK.giai_thich(period, _tt)
            return out
        if _tt == _DSK.MOI_TAO:
            out["ky_moi_tao"] = period
        # idempotent: xoá MỌI dòng cũ CÙNG source_file (mỗi file phủ trọn 1 tháng ngày).
        # KHÔNG giới hạn ở REPORT_TYPES nữa (sửa 2026-08-06): file báo cáo ngày của đơn vị CHƯA
        # khai trong `_UNITS` sẽ rơi xuống pipeline THÁNG (gate `is_daily_report` trả False) và để
        # lại dòng report_type THÁNG mang đúng source_file này — vd HT kỳ 2026-08 có 40 dòng
        # HQKD/PNLT/CHIPHI/DTHU sinh từ chính file ngày, hiện lên dashboard như số CẢ THÁNG trong
        # khi thực chất là vài ngày cộng lại. Khi đơn vị được thêm vào `_UNITS`, deriver này là chủ
        # DUY NHẤT của file (agent_cli return sớm ở gate) -> mọi dòng khác cùng source_file đều là
        # rác của lần phân loại nhầm trước đó, phải dọn cùng lượt nạp.
        # CHỐT 18/09/2026: xoá trần CHỈ khi khoá `source_file` thuộc về ĐÚNG MỘT file vật lý.
        # Khi có bản sinh đôi ở thư mục report_type khác (xem `_ban_sinh_doi`), hai bộ dòng đến từ
        # HAI file thật chứ không phải một file bị phân loại nhầm — thu hẹp về `_D` ở ĐÚNG ca này
        # nên KHÔNG mở lại lỗi 06/08 (ca đó chỉ có một file vật lý, nhánh trên vẫn xoá trần).
        if _ban_sinh_doi(path):
            _rt_ngay = tuple(REPORT_TYPES)
            cur.execute("DELETE FROM raw_rows WHERE source_file=%s AND report_type = ANY(%s)",
                        (source_file, list(_rt_ngay)))
        else:
            cur.execute("DELETE FROM raw_rows WHERE source_file=%s", (source_file,))
        if snap_mode:
            # Dọn nốt dòng mang khoá THEO FILE của cùng kỳ: mỗi snapshot từng đi qua pipeline
            # THÁNG (gate `is_daily_report` chỉ nhận khi đơn vị đã khai `_UNITS`) và để lại dòng
            # report_type tháng mang tên file đó — chính là số LUỸ KẾ hiện lên như số cả tháng.
            cur.execute("DELETE FROM raw_rows WHERE source_file LIKE %s AND source_file <> %s",
                        (f"{folder}::%.D.{period[:4]}{period[5:7]}%", source_file))
        payload = json.dumps({"unit": "ty", "grain": "day"}, ensure_ascii=False)
        recs, i = [], 0
        for ngay, facts in per_day:
            for f in facts:
                # Phần tử thứ 6 (dim2) là TUỲ CHỌN — layout "srvf" dùng để gắn kênh bán, layout
                # "xdv" dùng để gắn YẾU TỐ CHI PHÍ (6 nội dung, review DB 04/09/2026); còn lại
                # cụm bán xe. Giải nén theo lát cắt thay vì đổi mọi layout sang tuple 6 phần tử.
                cc, rt, dim1, dim3, v = f[:5]
                dim2 = f[5] if len(f) > 5 else None
                # Phần tử thứ 7 = payload RIÊNG của dòng (JSON đã dựng sẵn). Dòng SỐ DƯ cần mang
                # du_dau/ps_tang/ps_giam/ma_dt y như bản THÁNG, không dùng chung payload cả lượt.
                pl = f[6] if len(f) > 6 and f[6] else payload
                i += 1
                # GHI FULL PRECISION — `round(v * 1e-9, 9)` là làm tròn tới ĐỒNG (9 chữ số thập
                # phân của TỶ = 1 đồng), tức làm tròn TỪNG Ô trước khi dashboard cộng lại. Đúng
                # cái quy ước 30/07/2026 cấm: kế toán đối chiếu tới từng đồng, mà "làm tròn từng
                # phần rồi mới cộng" luôn lệch so với "cộng rồi làm tròn một lần".
                # Bắt được 04/09/2026 qua log kiểm soát Xanh Taxi: LNST khối luỹ kế 14 ngày, số
                # thật −1.674.139.469,54 -> làm tròn một lần ra −1.674.139.470 (đúng số kế toán),
                # còn cộng 84 ô đã tròn ra −1.674.139.468. File nguồn có phần lẻ thật (01/08 Phú
                # Thọ = −31.628.830,804), làm tròn từng ô là vứt phần lẻ đó đi.
                # Làm tròn chỉ ở TẦNG HIỂN THỊ. Các `round(sum(...) * 1e-9, 9)` ở phần tóm tắt
                # dry-run bên trên thì ĐÚNG: cộng trước, làm tròn một lần.
                recs.append((dataset_id, rt, 6200000 + i, ngay,
                             _CC_CONGTY.get(cc) or unit["cong_ty"], unit["khoi"], cc, period,
                             v * 1e-9, None, dim1, dim2, dim3, pl, source_file))
        cur.executemany(
            "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
            "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, source_file) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", recs)
        conn.commit()
        out["written"] = len(recs)
    finally:
        conn.close()
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--write", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    print(json.dumps(derive(a.file, a.write), ensure_ascii=False, indent=2))

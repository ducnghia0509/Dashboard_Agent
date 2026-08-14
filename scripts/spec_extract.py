# -*- coding: utf-8 -*-
"""ENGINE TRÍCH XUẤT THEO SPEC — thay cho việc viết 1 file Python cho mỗi nguồn báo cáo.

VÌ SAO CÓ FILE NÀY (2026-08-08): mọi deriver trước đây (`derive_hqkd_ngay.py`,
`derive_congno_tuoino.py`, `derive_vhkd_kqkd.py`…) đọc cột theo CHỈ SỐ CỨNG. Kế toán chèn/xoá 1 cột
là sai số im lặng, và mỗi lần mapping đổi là một vòng sửa Python → test → deploy. Đợt VHKD/XDV có
~10 nguồn nên chi phí đó nhân lên. Engine này đọc một **spec JSON** mô tả nguồn và ánh xạ cột, nên
đổi mapping = sửa JSON, không đụng code.

NGUYÊN TẮC SỐ 1 — DÒ CỘT THEO **TÊN HEADER**, KHÔNG THEO CHỮ CỘT.
Riêng điều này diệt cả một lớp lỗi đã gặp thật: file mapping ghi kho xe B2B "Trạng thái >30/45
ngày" ở cột AK, file thật để ở AM (AK là "Số ngày tồn kho thực tế"); kho xe B2C ghi "Tên KH" cột O,
thật ra ở P. Có thể khai `"cot_du_phong": "AM"` làm phao khi header rỗng, nhưng tên luôn thắng.

SPEC ĐẶT Ở: `Dashboard_Agent/extract_specs/<id>.json` (KHÔNG để trong `memory/` — thư mục đó
bị .gitignore vì là state runtime của agent; spec là CẤU HÌNH NGUỒN, phải được version). Đây là file CẤU HÌNH TIN CẬY của nội
bộ (không nhận từ ngoài) — `dan_xuat` dùng eval trong sandbox hẹp; đừng mở cho input người dùng.

CẤU TRÚC SPEC (khoá tiếng Việt cho kế toán/BA đọc được):
{
  "id": "vhkd_kqkd",                     // trùng tên file
  "report_type": "KDVH",                 // ghi vào raw_rows.report_type
  "row_index_base": 6400000,             // dải row_index riêng, tránh đụng deriver khác
  "khoi": "Khối KD Vinfast - Showroom",  // hằng số (hoặc bỏ, nếu suy từ cost_center)
  "nguon": {
    "folder": "SRVF/baocaokqkd",         // dưới received_reports/
    "file_glob": "*.xlsx",
    "sheet": {"batdau": "CHI TIẾT XHĐ"}  // | {"ten": "..."} | {"chua": "..."} | {"so": 0}
                                        // | {"theo_thang": "T{mm}"} — file 1 sheet/tháng (BCTC SRVF)
  },
  "header": {"dong": 1},                 // | {"dong": [1,2]} gộp 2 dòng header (lấy ô đầu khác rỗng)
                                         // | {"tim_o": "Mã số"} tự dò dòng header theo nhãn mốc
                                         //   (dùng khi vị trí header khác nhau giữa các sheet)
  "dong_bat_dau": 3,                     // tuỳ chọn — mặc định = dòng header cuối + 1
  "dong_ket_thuc": 19,                   // tuỳ chọn — chặn trên của dải dòng (bắt buộc khi dùng `vung`)
  "nam_tu_ten_file": {"regex": "\\.M\\.(\\d{4})\\."},   // cho `kieu: "thang_cuoi"`
  "vung": [                              // NHIỀU KHỐI trong CÙNG 1 sheet (xem `extract_file`).
    {"ten": "Sơn Tây", "header": {"dong": [6, 7]}, "dong_bat_dau": 8, "dong_ket_thuc": 19,
     "chieu_co_dinh": {"cost_center": "ST_AT"}, "cot": {"amount": {"header": "TXTX"}}}
  ],                                     // khoá trong vùng ghi đè spec; `cot`/`chieu_co_dinh` trộn
  "chieu_tu_ten_file": {"dim2": {"regex": "Xuathoadon_(B2B|B2C|GF)", "hoa": true}},
  "ngay_tu_ten_file": {"regex": "M\\.(\\d{4})\\.(\\d{1,2})\\.(\\d{1,2})", "thu_tu": "ymd"},
  "cot": {                               // đích -> cách lấy. Đích: ngay/cost_center/cong_ty/
    "cost_center": {"header": "Tên DVCS", "chuan_hoa": "sr_showroom"},   // amount/amount2/dim1..3/
    "ngay":   {"header": "Ngày hóa đơn", "kieu": "date"},                // payload.<khoá bất kỳ>
    "amount": {"header": "Giá bán", "kieu": "so", "he_so": 1e-9}
  },                                     // `kieu` khác: "thang_cuoi" (ô là SỐ THÁNG 1..12, năm
                                         // lấy từ tên file -> ngày cuối tháng) · "ngay_trong_thang"
                                         // (ô là SỐ NGÀY, năm+tháng từ tên file). Hai kiểu này cho
                                         // báo cáo xếp mỗi kỳ MỘT DÒNG thay vì một cột.
  "ban_ghi": "moi_dong",                 // | "moi_cot_gia_tri" | "moi_cot_ngay"
                                         // | "moi_cot_thang" (xem dưới)
  "cot_gia_tri": [                       // chỉ dùng khi ban_ghi = "moi_cot_gia_tri":
    {"header": "Công nợ trong hạn", "dim1": "Trong hạn", "he_so": 1e-9}
  ],                                     // -> mỗi dòng nguồn đẻ N bản ghi, amount lấy từng cột
  "loc": [{"cot": "ngay", "dieu_kien": "khac_rong"}],
  "dan_xuat": {"payload.lng": "amount - payload.gia_von"},
  "payload_them": {"unit": "ty"}
}

`ban_ghi`:
  - "moi_dong"        : 1 dòng nguồn -> 1 bản ghi (bảng chi tiết: xuất hoá đơn, claim, kho xe).
  - "moi_cot_gia_tri" : 1 dòng nguồn -> N bản ghi, mỗi cột giá trị thành 1 dòng có dim1 riêng
                        (bảng ma trận: "Tuổi nợ phải thu" có sẵn cột trong hạn/1-30/>30-90/…).
  - "moi_cot_ngay"    : 1 dòng nguồn -> N bản ghi theo dải cột NGÀY trong tháng (`cot_ngay`).
  - "moi_cot_thang"   : 1 dòng nguồn -> N bản ghi theo dải cột THÁNG (`cot_thang`), mỗi bản ghi
                        neo vào ngày cuối tháng đó. Dùng cho bản KẾ HOẠCH năm: một file duy nhất
                        cấp số cho cả 12 kỳ, nên đừng đặt kỳ theo tên file.

Ghi DB idempotent theo `source_file` = "<FOLDER>::<tên file>" — cùng quy ước mọi deriver khác.
`--write` mới ghi; mặc định dry-run in ra tổng hợp để đối chiếu với file gốc.

HOOK (khi spec không tả nổi): thêm hàm vào `_CHUAN_HOA` bên dưới rồi gọi bằng tên trong
`"chuan_hoa"`. Hook trả str, hoặc dict để set nhiều trường cùng lúc (vd cost_center + cong_ty).
"""
import argparse
import calendar
import datetime as dt
import fnmatch
import glob
import json
import os
import re
import sys
import unicodedata

import openpyxl
import psycopg
from openpyxl.utils import column_index_from_string, get_column_letter

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(_ROOT, ".env"))

DB_URL = os.environ.get("DATABASE_URL")
SPEC_DIR = os.path.join(_ROOT, "extract_specs")
REPORTS_DIR = os.path.normpath(os.path.join(_ROOT, "..", "Connect_VPS", "received_reports"))

_CHUAN_TRUONG = {"ngay", "cong_ty", "khoi", "cost_center", "amount", "amount2",
                 "dim1", "dim2", "dim3"}


# ─────────────────────────── tiện ích chuẩn hoá ───────────────────────────
def _nd(s):
    """Bỏ dấu + bỏ ký tự không phải chữ/số + thường hoá — dùng để so tên header/tên đơn vị."""
    s = str(s if s is not None else "").strip().lower().replace("đ", "d")
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s)


# Giá trị LỖI của Excel: VLOOKUP chưa tra được, ô tham chiếu bị xoá, chia 0… File nguồn vẫn gửi
# về nguyên trạng chuỗi này. Coi như Ô TRỐNG ở MỌI spec — nếu không, chuỗi "#N/A" đi thẳng vào
# dim/cost_center và trở thành một "giá trị nghiệp vụ" giả: đã dính 14/08/2026 ở claim B2C T8
# (cột Trạng thái 178 ô #N/A, cột Số tiền hỏng CẢ 605 ô) -> dashboard đếm 175 hồ sơ trạng thái
# "#N/A" giá trị 0đ và đòi bổ sung "#N/A" vào bảng quy ước trạng thái claim.
# Số ô lỗi được ĐẾM và báo ra `canh_bao` (xem cuối `extract_file`): ô lỗi là dấu hiệu file nguồn
# chưa cập nhật xong, phải nhìn thấy chứ không nuốt im lặng.
_LOI_EXCEL = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#GETTING_DATA"}


def _o_loi(v):
    return isinstance(v, str) and v.strip().upper() in _LOI_EXCEL


# Câu cảnh báo "đã đọc được file nhưng bộ lọc loại hết dòng" — `run()` dựa vào nó để phân biệt
# "file hết số" với "file không đọc được" (xem chú thích ở `run`). Một chỗ khai, hai chỗ dùng.
_W_BO_LOC = "dòng không qua bộ lọc"


def _so(v, he_so=1.0):
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) * he_so
    s = str(v or "").strip().replace(" ", "")
    if not s:
        return 0.0
    # '1.234.567,89' (VN) và '1,234,567.89' (EN) đều gặp trong file kế toán
    s = s.replace(".", "").replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        return float(s) * he_so
    except ValueError:
        return 0.0


def _so_co_rong(v, he_so=1.0):
    """Như `_so` nhưng ô TRỐNG/không đọc được -> None, KHÔNG phải 0.0.

    Dùng cho payload mà "0" là một giá trị có nghĩa thật, khác hẳn "chưa biết". Ca cụ thể (14/08/
    2026): 'Số ngày tồn thực tế' của file tồn kho GF — 12/62 xe ghi 'Ngày xe về SR' = 'Chưa' nên ô
    tuổi tồn là #VALUE!. Qua `_so` thì thành 0.0, tức 12 xe đó đọc thành "mới về hôm nay, tồn 0
    ngày" và lọt vào nhóm dưới 30 ngày -> tỷ lệ quá hạn TỰ HẠ, không cảnh báo gì. Còn xe về đúng
    ngày chốt thì 0 là số THẬT, nên không thể lấy `if not v` mà suy ra "chưa biết".
    ĐỪNG dùng cho `amount`: tầng đọc số cộng tiền bằng `r["amount"] or 0`, None ở đó vô hại nhưng
    cũng không thêm thông tin gì; giữ `so` cho amount để không đổi hành vi các spec đang chạy.
    """
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) * he_so
    s = str(v).strip().replace(" ", "")
    s = s.replace(".", "").replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        return float(s) * he_so
    except ValueError:
        return None


def _date(v):
    """-> 'YYYY-MM-DD' hoặc None. File có cả datetime, 'dd/mm/yyyy' lẫn 'yyyy-mm-dd'."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        d, mo, y = (int(x) for x in m.groups())
    else:
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if not m:
            return None
        y, mo, d = (int(x) for x in m.groups())
    try:
        return dt.date(y, mo, d).isoformat()
    except ValueError:
        return None


# ─────────────────────────── hook chuẩn hoá ───────────────────────────
_CC_CACHE = {}


def _cc_showroom(ten):
    """Tên đơn vị trong file -> {cost_center, cong_ty} của khối Vinfast - Showroom.

    File ghi "Showroom OceanPark"/"Showroom Uông Bí", master ghi "Vinfast Ocean Park" -> bỏ tiền
    tố showroom/vinfast ở CẢ hai phía rồi so. Bản "(61)" (pháp nhân Xanh Vĩnh Phúc) chuẩn hoá
    thành "...61" nên không đụng bản gốc. "Vinfast B2B" là đội bán B2B tập trung -> B2B_SR.
    """
    return _cc_theo_khoi(ten, "Khối KD Vinfast - Showroom", {"b2b": ("B2B_SR", "TC")})


def _cc_hcns_xdv(ten):
    """Phòng/Ban của báo cáo NHÂN SỰ -> cost center XDV, CHỈ khi tên đó THỰC SỰ là một XƯỞNG.

    VÌ SAO KHÔNG DÙNG THẲNG `_cc_xdv` (bắt được 14/08/2026): `_bo_tien_to` bóc cả tiền tố
    "showroom", nên "Showroom Vinfast Smart City" quy về cùng khoá "smartcity" với xưởng
    "Vinfast Smart City" -> CẢ 9 SHOWROOM bị gán cost center XDV, cộng thêm 470 nhân sự BÁN HÀNG
    vào mẫu số "nhân sự xưởng" (994 thay vì 524, phồng 90%). Các spec XDV khác không dính vì file
    của chúng chỉ có xưởng; báo cáo nhân sự thì chứa MỌI khối trong một sheet.

    Báo cáo nhân sự gọi xưởng bằng đúng hai dạng: "Xưởng dịch vụ ..." và "Trung tâm Sửa chữa Pin,
    Động cơ ..." (xưởng HCM). Chốt bằng tiền tố đó là tách được Showroom khỏi xưởng.

    CÒN MỘT CA KHÔNG CHỐT ĐƯỢC Ở ĐÂY: "Xưởng dịch vụ Hà Khánh" xuất hiện ở CẢ khối "Khối Dịch vụ
    Hậu mãi" (33 người) và "Khối Kinh doanh Showroom Vinfast" (3 người) — hook chỉ thấy MỘT ô nên
    không biết dòng đang thuộc khối nào. Bên đọc PHẢI lọc thêm `dim1 = "Khối Dịch vụ Hậu mãi"`
    (xem `app/metrics/xdv.py::_nhan_su_xuong`), nếu không cộng đôi 3 người.
    """
    n = _nd(ten)
    if not (n.startswith("xuongdichvu") or n.startswith("trungtamsuachua")):
        return {"_khong_map": str(ten or "").strip()}
    return _cc_xdv(ten)


def _cc_xdv(ten):
    """Tên xưởng trong file -> cost center khối XDV. 13/14 xưởng khớp thẳng sau khi bỏ tiền tố
    ("Ocean Park" ↔ master "Vinfast Ocean Park"); riêng xưởng HCM có tới BA tên gọi nên phải
    khai alias: master "Vinfast Hồ Chí Minh", bản THỰC HIỆN ghi "HCM", bản KẾ HOẠCH
    (Baocaodoanhthukehoachngay, 10/08/2026) ghi theo quận — "Quận 12"."""
    # Khoá alias phải viết theo dạng ĐÃ CHUẨN HOÁ của `_nd` — bỏ dấu, bỏ cả khoảng trắng:
    # "Quận 12" -> "quan12". Viết "quan 12" thì không bao giờ khớp và xưởng đó lặng lẽ mất.
    # Alias thứ BA cho xưởng HCM (14/08/2026): báo cáo NHÂN SỰ (HCNS/baocaotongsonhansu, sheet
    # "Chi tiết") gọi nó là "Trung tâm Sửa chữa Pin, Động cơ Thành phố Hồ Chí Minh". Không khai
    # thì đúng 1 trong 14 xưởng rơi và mẫu số "nhân sự xưởng" hụt 35 người (~6,7%) — kiểu thiếu
    # vừa đủ nhỏ để không ai thấy.
    return _cc_theo_khoi(ten, "Khối KD Vinfast - XDV",
                         {"hcm": ("HCM_XDV", "TC"), "quan12": ("HCM_XDV", "TC"),
                          "trungtamsuachuapindongcothanhphohochiminh": ("HCM_XDV", "TC")})


def _master_loader():
    """Nạp `app/master_data/loader.py` của backend THEO ĐƯỜNG DẪN FILE, không theo tên package.

    Vì sao không `sys.path.insert(...)` rồi `from app.master_data import loader` như trước: khi
    hàm này chạy TRONG tiến trình đã import `servers/template_filler.py` (đường "Đồng bộ & nạp" và
    bước nạp lại của theo-dõi-thay-đổi), `app` ĐÃ nằm trong sys.modules và trỏ vào backend CŨ
    (`DashBoard_AI/backend`, theo BACKEND_PATH) — bản đó không có `master_data`, nên import ném
    ModuleNotFoundError và toàn bộ spec có chuẩn hoá cost center trả 0 dòng. Chạy CLI thì lại
    không lỗi (chưa ai import `app`), nên bug chỉ hiện khi bấm nút trên web — đúng kiểu khó tìm.
    Nạp theo đường dẫn file thì tên `app` của ai cũng không ảnh hưởng.

    loader.py chỉ dùng stdlib và đọc JSON trong `app/data/` nên nạp rời hoàn toàn an toàn.
    """
    import importlib.util
    ung_vien = [
        os.environ.get("MASTER_DATA_BACKEND"),
        os.path.join(_ROOT, "..", "AI_coding", "tc-admin-api"),      # bản test
        os.path.join(os.path.expanduser("~"), "apps/tc-console/tc-admin-api"),   # bản prod
    ]
    for base in ung_vien:
        if not base:
            continue
        p = os.path.normpath(os.path.join(base, "app", "master_data", "loader.py"))
        if not os.path.isfile(p):
            continue
        spec = importlib.util.spec_from_file_location("tc_master_data_loader", p)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    # Đường cũ làm lưới cuối: có thể chạy ở nơi bố cục thư mục khác hẳn dự đoán ở trên.
    sys.path.insert(0, os.path.join(_ROOT, "..", "AI_coding", "tc-admin-api"))
    from app.master_data import loader as master   # noqa: PLC0415
    return master


def _cc_theo_khoi(ten, khoi, alias):
    if khoi not in _CC_CACHE:
        master = _master_loader()
        m = {}
        for cc in master.master_data().get("costCenters", []):
            if (cc.get("khoi") or "") != khoi:
                continue
            m[_bo_tien_to(_nd(cc.get("ten")))] = (str(cc.get("ma") or "").strip(),
                                                 master.resolve_company_code(cc.get("congTy") or ""))
        _CC_CACHE[khoi] = m
    key = _bo_tien_to(_nd(ten))
    ma, cty = _CC_CACHE[khoi].get(key) or alias.get(key) or (None, None)
    # TRẢ LUÔN `khoi` (14/08/2026). Trước đây chỉ trả cost_center + cong_ty, và mọi spec dùng hook
    # này đều tự khai `"khoi"` ở cấp spec nên không ai thấy thiếu. Nhưng báo cáo NHÂN SỰ chứa NHIỀU
    # khối trong một sheet -> không khai được `khoi` cố định, mà thiếu `khoi` thì `filter_scope`
    # (màn xdv0 tự lọc khối) LỌC SẠCH các dòng đó và mẫu số "nhân sự xưởng" về None mà không có lỗi
    # nào. Giá trị trả về TRÙNG KHỚP với `khoi` mà các spec cũ vẫn khai, nên chúng không đổi hành vi.
    return ({"cost_center": ma, "cong_ty": cty, "khoi": khoi} if ma
            else {"_khong_map": str(ten or "").strip()})


def _bo_tien_to(n):
    # "thinhcuong": bản công nợ tuần 10/08/2026 đổi cách ghi tên xưởng, từ "Ocean Park" sang
    # "Vinfast Thịnh Cường Ocean Park" -> bỏ "vinfast" xong vẫn còn "thinhcuong" nên KHÔNG khớp
    # danh mục, cả 14 xưởng rơi hết và file nạp ra 0 dòng. Không cost center nào của master bắt
    # đầu bằng "Thịnh Cường" (mấy dòng Xe tải để ở ĐUÔI) nên bỏ tiền tố này là an toàn.
    # BÓC LẶP, không phải một lượt (sửa 14/08/2026). Vòng `for` một lượt xét tiền tố theo THỨ TỰ
    # khai, nên "Xưởng dịch vụ Vinfast Cẩm Phả" bóc được "xuongdichvu" rồi DỪNG, còn lại
    # "vinfastcampha" trong khi khoá master là "campha" -> 13/14 xưởng của báo cáo nhân sự HCNS
    # rơi hết. Bóc lặp thì cả hai vế (tên trong file và tên master) cùng quy về một dạng, và vì
    # danh mục cost center được chuẩn hoá bằng CHÍNH hàm này nên không sinh khoá lệch.
    doi = True
    while doi:
        doi = False
        for p in ("showroom", "vinfast", "thinhcuong", "xuongdichvu", "xdv"):
            if n.startswith(p) and len(n) > len(p):    # giữ nguyên nếu bóc xong thành rỗng
                n, doi = n[len(p):], True
    return n


# Nhãn dòng của sheet KHDT (bản kế hoạch) -> (mã nhóm, kênh). Khớp CHÍNH XÁC sau chuẩn hoá, KHÔNG
# khớp kiểu "chứa": ngay dưới các dòng chi tiết còn có dòng tổng "Tổng doanh thu kênh B2C" —
# khớp kiểu chứa là cộng đôi toàn bộ kế hoạch mà không có dấu hiệu gì.
# Mã nhóm theo ĐÚNG file kế hoạch: A230 = doanh thu khác, A250 = claim (thiết kế
# Dashboard.dc.html ghi ngược hai mã này — đã chốt với nghiệp vụ 2026-08-10, lấy theo file).
_KH_DONG = {}
for _ma, _nhan in (("A200", "Bán xe kênh {}"),
                   ("A230", "Doanh thu khác Kênh {}"),
                   ("A250", "Claim kênh {}")):
    for _k in ("B2C", "B2B", "GF"):
        _KH_DONG[_nd(_nhan.format(_k))] = (_ma, _k)


def _kh_dong(nhan):
    """Nhãn dòng kế hoạch -> {dim1: mã nhóm, dim2: kênh}. Dòng tổng/tiêu đề -> không đặt gì
    (bản ghi thiếu dim2 sẽ bị `loc` loại), nên chỉ các dòng CHI TIẾT theo kênh được giữ."""
    got = _KH_DONG.get(_nd(nhan))
    return {"dim1": got[0], "dim2": got[1]} if got else {}


# Dòng CHỈ TIÊU TỔNG của bản kế hoạch XDV (sheet "KHDT XHD" mục IV, V): nhãn nằm NGAY ở cột B của
# chính dòng có số, không có dòng tiêu đề riêng như mục I.1/I.2/II/III — nên không dùng
# `ngu_canh_dong` được (dòng tiêu đề không tự sinh bản ghi).
# THỨ TỰ QUAN TRỌNG: "kehoachtongchiphi" là TIỀN TỐ của "kehoachtongchiphivanhanh…", so kiểu
# "chứa" mà xét CP_TONG trước là mục V bị gán nhầm thành CP_TONG rồi hai dòng đè nhau.
_XDV_KH_CHI_TIEU = (
    ("kehoachtongchiphivanhanh", "CP_VH"),
    ("kehoachtongchiphi", "CP_TONG"),
)


def _xdv_kh_dong(nhan):
    """Cột B bản kế hoạch XDV -> chỉ tiêu tổng (mục IV/V) HOẶC xưởng.

    `dim2` là cờ phân tầng, BẮT BUỘC có để `loc` giữ được đúng dòng thật:
      · "Tổng"  — dòng cấp khối (mục IV/V, và dòng "Cộng" bắt ở cột A bởi `_xdv_kh_cong`)
      · "Xưởng" — dòng chi tiết theo cost center
    Không có cờ này thì mấy ô rác lạc giữa sheet (vd ô "1" ở H45 sheet "KHDT RO") vẫn mang dim1
    thừa hưởng từ ngữ cảnh mục đang mở và lọt vào DB thành một dòng kế hoạch vô nghĩa.
    """
    n = _nd(nhan)
    for key, ma in _XDV_KH_CHI_TIEU:
        if key in n:
            return {"dim1": ma, "dim2": "Tổng"}
    res = _cc_xdv(nhan)
    if res.get("cost_center"):
        res["dim2"] = "Xưởng"
    return res


# Sheet "CT Vận hành" của bản kế hoạch XDV — NGƯỠNG MỤC TIÊU vận hành T7..T12, nhãn ở cột B.
# CHỈ khai 5 chỉ tiêu KHÔNG trùng nguồn khác. 6 dòng đầu sheet (doanh thu, sản lượng, lợi nhuận
# gộp, LNST, doanh số & sản lượng lệnh nghiệm thu) lặp lại y hệt số của "KHDT XHD"/"KHDT RO" đã
# nạp thành XDV_KH*/XDV_KH_RO* — nạp thêm là tạo hai nguồn cho cùng một con số, sớm muộn lệch nhau.
# Khớp kiểu "chứa" sau `_nd`, nên chịu được cả lỗi gõ "Gía trị" lẫn đuôi giải thích dài trong ngoặc.
_XDV_CT_DONG = (
    ("giatrilenhnghiemthuchuaxuathoadon", "RO_CHUA_XHD"),
    ("lenhnghiemthuchuaxuathoadon30ngay", "RO_CHUA_XHD_30"),
    ("lenhsuachua10ngaychuanghiemthu", "LSC_10N"),
    ("giatribinhquandonhang", "BQ_DON"),
    ("nangsuatld", "NSLD"),
)


def _xdv_ct_dong(nhan):
    """Cột B sheet "CT Vận hành" -> mã ngưỡng. Dòng không khai -> {} (giữ nguyên, `loc` sẽ loại).

    Trả `{}` chứ KHÔNG trả None: hook trả dict thì engine gộp khoá rồi bỏ qua, nên dòng lạ không
    bị ghi đè dim1 thành None và cũng không sinh cảnh báo "không map được" vô nghĩa.
    """
    n = _nd(nhan)
    for key, ma in _XDV_CT_DONG:
        if key in n:
            return {"dim1": ma, "dim2": "Tổng"}
    return {}


# Sheet "CT Vận hành" của bản kế hoạch SHOWROOM (`1.SR.*.Kehoachthang.xlsx`) — KHÁC HẲN sheet
# cùng tên của bản XDV, đừng dùng lại `_xdv_ct_dong`: bản SR không có cột "Công thức tính" nên dải
# tháng lùi một cột (E..J thay vì F..K), và 10 nhãn dòng hoàn toàn khác.
# CHỈ khai 2 dòng NĂNG SUẤT. Tám dòng còn lại cố ý bỏ, mỗi dòng một lý do:
#   - "Doanh thu" (CT33) và "Sản lượng" (CT34) là công thức `=KHDT!F16` / `=KHDT!F21` — trỏ thẳng
#     vào sheet KHDT đã nạp thành VHKD_KH/VHKD_KH_SL. Đã đối chiếu T7: CT33 = 823.271.212.048 =
#     đúng tổng A200 ba kênh, CT34 = 1.414 = đúng tổng sản lượng ba kênh. Nạp thêm là hai nguồn
#     cho một con số.
#   - "Gía trị bình quân đơn hàng" (CT36) là `=D5/D7`, `app/metrics/vhkd.py` đã tự suy ra.
#   - "Lợi nhuận" bỏ trống toàn bộ trong file.
#   - 4 dòng cuối (xe tồn > 30 ngày, claim chưa xác nhận, claim chưa làm hồ sơ, công nợ quá hạn
#     đã có COC) ghi ngưỡng BẰNG LỜI ("Nhỏ hơn 30% tồn kho", "Bằng 0 - cứ có 1 khách hàng quá
#     hạn, trừ 5% mức độ hoàn thành") -> `_so` ra 0 nên `moi_cot_thang` tự bỏ. CHỦ Ý: muốn chấm
#     %HT cho 4 nhóm đó thì phải chốt ngưỡng thành SỐ với nghiệp vụ trước, không phải đoán ở đây.
# MÃ Ở CỘT A CỦA FILE ĐỀ LỆCH MAPPING — bám nhãn dòng, không bám mã: mapping ghi năng suất sản
# lượng = CT37 và giá trị BQ đơn hàng = CT38, còn file ghi CT38 và CT36. `app/metrics/vhkd.py`
# đang theo mã của FILE, nên hook cũng gắn mã theo file.
_SR_CT_DONG = (
    ("nangsuatdoanhthu", "NS_DT"),
    ("nangsuatsanluong", "NS_SL"),
)


def _sr_ct_dong(nhan):
    """Cột B sheet "CT Vận hành" bản Showroom -> mã chỉ tiêu năng suất. Dòng khác -> {}.

    Khớp kiểu "chứa" sau `_nd` để chịu được dấu cách thừa của file ("Năng suất Sản lượng /1NV").
    Thứ tự trong `_SR_CT_DONG` không quan trọng ở đây vì hai khoá không lồng nhau, nhưng ĐỪNG
    thêm khoá ngắn kiểu "sanluong": nó nằm trong "nangsuatsanluong1nvbanhang" và sẽ bắt sai dòng.
    """
    n = _nd(nhan)
    for key, ma in _SR_CT_DONG:
        if key in n:
            return {"dim1": ma, "dim2": "Tổng"}
    return {}


def _xdv_kh_cong(v):
    """Cột A -> đánh dấu dòng "Cộng" (tổng khối) của từng mục.

    Dòng tổng của bản kế hoạch XDV ghi chữ "Cộng" ở cột A và BỎ TRỐNG cột B (khác hẳn các dòng
    xưởng). Mục III (lợi nhuận sau thuế) CHỈ có đúng dòng này — bỏ qua là mất trắng cả mục.
    """
    return {"dim2": "Tổng"} if _nd(v) in ("cong", "tong", "tongcong") else {}


_CHUAN_HOA = {
    "sr_showroom": _cc_showroom,
    "xdv": _cc_xdv,
    "hcns_xdv": _cc_hcns_xdv,
    "kh_dong": _kh_dong,
    "xdv_kh_dong": _xdv_kh_dong,
    "xdv_kh_cong": _xdv_kh_cong,
    "xdv_ct_dong": _xdv_ct_dong,
    "sr_ct_dong": _sr_ct_dong,
    "hoa": lambda v: str(v or "").strip().upper() or None,
    "cat": lambda v: str(v or "").strip() or None,
}


# ─────────────────────────── đọc spec & sheet ───────────────────────────
def load_spec(ref):
    path = ref if os.path.isfile(ref) else os.path.join(SPEC_DIR, f"{ref}.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def specs_for_path(path):
    """Các spec phụ trách FILE này, suy từ `nguon.folder` — [] nếu file không thuộc nguồn nào.

    Vì sao cần: `agent_cli.cmd_autofill` (nút "Phân tích AI") và `template_filler.autofill_file`
    ("Đồng bộ & nạp" + bước nạp lại của theo-dõi-thay-đổi) là hai cửa DUY NHẤT mọi nguồn đi qua,
    và cả hai điều phối bằng một bảng viết tay: mỗi loại báo cáo một deriver. 18 nguồn khai bằng
    JSON spec ra sau nên chưa có tên trong bảng đó — bấm ingest thì báo thành công mà 0 dòng, im
    lặng, không ai hiểu vì sao (đúng lỗi user gặp 10/08/2026). Hàm này là mảnh nối còn thiếu.

    MỘT THƯ MỤC CÓ THỂ CÓ NHIỀU SPEC: `XDV/baocaodoanhthungay` cấp cho 3 spec (doanh thu vụ
    việc / doanh thu RO / số lượng RO) và `KEHOACH/baocaokehoachthang` cấp cho 2. Phải trả về
    TẤT CẢ, chạy thiếu một cái là màn hình thiếu đúng một khối số.
    """
    p = os.path.normpath(os.path.abspath(path)).replace("\\", "/")
    out = []
    for f in sorted(glob.glob(os.path.join(SPEC_DIR, "*.json"))):
        try:
            with open(f, encoding="utf-8") as fh:
                sp = json.load(fh)
        except Exception:
            continue                      # spec hỏng cú pháp KHÔNG được làm chết luồng nạp
        folder = ((sp.get("nguon") or {}).get("folder") or "").strip("/")
        if not folder:
            continue
        if f"/{folder}/" in p:
            out.append(sp)
    return out


def spec_doc_tron_file(path) -> bool:
    """Các spec của file này có ĐỌC TRỌN file không, hay chỉ bóc một lát của nó?

    True  -> caller được phép ra sớm: spec là nguồn DUY NHẤT của file (mọi nguồn VHKD/XDV, file
             sinh ra chỉ để cấp đúng bộ chỉ tiêu mà spec tả).
    False -> file còn cấp report_type khác qua đường tất định (deriver CĐKT/CĐPS/tồn kho/công nợ…),
             caller PHẢI chạy tiếp sau khi spec xong, không được return.

    Đánh dấu bằng `doc_tron_file: false` trong chính spec — mặc định True để 26 spec hiện có giữ
    NGUYÊN hành vi cũ, chỉ spec nào tự khai "tôi chỉ bóc một lát" mới đổi. Suy đoán tự động (vd
    "file có sheet CĐKT thì chắc là BCTC") nghe tiện nhưng sai ở cả hai chiều và không ai đọc code
    ra được vì sao một file lại đi hai đường; khai tường minh trong spec thì đọc spec là biết.

    Xem agent_cli._cmd_autofill_impl (ca SRVF T05/T06 mất 13 report_type, 14/08/2026).
    """
    specs = specs_for_path(path)
    return all(sp.get("doc_tron_file", True) for sp in specs) if specs else True


def run_for_path(path, write=False):
    """Chạy mọi spec phụ trách file này. Trả [] nếu KHÔNG spec nào phụ trách -> caller đi đường cũ.

    Từng spec bọc riêng try/except: một spec lỗi không được kéo theo các spec còn lại, và tuyệt
    đối không được ném ra ngoài — hàm này nằm trên đường nạp CHUNG của mọi báo cáo.
    """
    ket_qua = []
    for sp in specs_for_path(path):
        # BỎ QUA BẢN ĐÃ BỊ THAY THẾ. Hàm này nạp ĐÚNG file được đưa vào, không nhìn sang các file
        # khác cùng kỳ — nên "Nạp lại tất cả" (sync_orchestrator reprocess) sẽ ghi lại dòng của
        # bản CŨ mà lượt quét cả thư mục đã cố tình loại, và kỳ đó lập tức đếm đôi trở lại. Đúng
        # 3.029 dòng claim T3-T6 vừa phải xoá tay ngày 12/08/2026 sẽ sống dậy chỉ bằng một cú bấm.
        if sp.get("moi_ky_lay_file_moi_nhat"):
            try:
                giu, _w = quet_nguon(sp)
                giu, _bo = loc_file_moi_nhat(sp, giu)
                if os.path.abspath(path) not in {os.path.abspath(x) for x in giu}:
                    ket_qua.append({"file": os.path.basename(path), "dong": 0,
                                    "spec": sp.get("id"), "report_type": sp.get("report_type"),
                                    "canh_bao": ["BỎ QUA — đã có bản MỚI HƠN của cùng kỳ; nạp bản "
                                                 "này vào là đếm đôi. Xoá tài liệu của nó thay vì nạp."]})
                    continue
            except Exception:                                      # noqa: BLE001
                pass          # không kiểm được thì cứ chạy như cũ, đừng chặn đường nạp
        try:
            r = run(sp, path, write=write)
        except Exception as e:                                     # noqa: BLE001
            r = {"file": os.path.basename(path), "dong": 0,
                 "canh_bao": [f"LỖI spec {sp.get('id')}: {type(e).__name__}: {e}"]}
        r["spec"] = sp.get("id")
        r["report_type"] = sp.get("report_type")
        ket_qua.append(r)
    return ket_qua


def _chon_sheet(wb, cfg, thang=None):
    """Chọn sheet theo cfg. `theo_thang` (vd "T{mm}") dành cho file có MỘT SHEET MỖI THÁNG.

    BCTC của SRVF là ca đó: cùng một file tháng 7 chứa cả T01..T07, mỗi sheet là một tháng. Lấy
    sheet đầu hay khớp theo tiền tố "T" đều sai — "T01." và "T01" chuẩn hoá về cùng chuỗi, còn
    "T01BC"/"TSCD" cũng bắt đầu bằng T. Phải neo theo ĐÚNG tháng của kỳ đang nạp và khớp CHÍNH XÁC.
    """
    cfg = cfg or {}
    if "so" in cfg:
        return wb.sheetnames[int(cfg["so"])]
    if "theo_thang" in cfg:
        if not thang:
            return None
        muon = _nd(str(cfg["theo_thang"]).replace("{mm}", f"{int(thang):02d}"))
        # ƯU TIÊN khớp tuyệt đối tên gốc trước, vì "T01." và "T01" cùng _nd() -> "t01".
        for name in wb.sheetnames:
            if str(name).strip() == str(cfg["theo_thang"]).replace("{mm}", f"{int(thang):02d}"):
                return name
        for name in wb.sheetnames:
            if _nd(name) == muon:
                return name
        return None
    for name in wb.sheetnames:
        n = _nd(name)
        if "ten" in cfg and n == _nd(cfg["ten"]):
            return name
        if "batdau" in cfg and n.startswith(_nd(cfg["batdau"])):
            return name
        if "chua" in cfg and _nd(cfg["chua"]) in n:
            return name
    return None


def _map_header(rows, dong):
    """{tên header đã chuẩn hoá: chỉ số cột 0-based}. `dong` là số hoặc list số (1-based).

    Gộp nhiều dòng: lấy ô KHÁC RỖNG ĐẦU TIÊN theo thứ tự khai — file Claim để nhóm cột ở dòng 1
    ("XE"/"PIN") còn tên cột thật ở dòng 2 (Z "SR", AA "Trạng thái", AB "Số tiền").
    """
    dongs = dong if isinstance(dong, list) else [dong]
    ncol = max((len(rows[d - 1]) for d in dongs if d - 1 < len(rows)), default=0)
    out, theo_dong = {}, {d: {} for d in dongs}
    for j in range(ncol):
        xong = False
        for d in dongs:
            r = rows[d - 1] if d - 1 < len(rows) else ()
            v = r[j] if j < len(r) else None
            n = _nd(v)
            if n:
                theo_dong[d].setdefault(n, j)
                if not xong:
                    out.setdefault(n, j)
                    xong = True
    return out, theo_dong


def _resolve_cot(spec, hmap, warn):
    """{đích: (chỉ số cột, cfg)} — dò theo TÊN, phao là `cot_du_phong` (chữ cột)."""
    out = {}
    for dich, cfg in (spec.get("cot") or {}).items():
        j = _tim_cot(hmap, cfg, dich, warn)
        if j is None:
            continue
        out[dich] = (j, cfg)
    return out


def _tim_cot(hmaps, cfg, nhan, warn):
    """Chỉ số cột cho 1 khai báo. `header` nhận CHUỖI hoặc LIST tên ứng viên (thử theo thứ tự).

    Vì sao cần list: cùng một cột nhưng file đổi tên giữa các kỳ — báo cáo Claim ghi "Check" ở
    T1/T2 rồi đổi thành "Trạng thái" từ T3. Khai `["Check", "Trạng thái"]` chạy được cả 6 file.

    Vì sao cần `dong_header`: T1/T2 còn có một cột MỒI ở dòng 1 cũng tên "Trạng thái" nhưng chứa
    ĐỊA CHỈ khách. Không giới hạn dòng header thì trạng thái claim biến thành "Thôn Xám, Xã Vĩnh
    Phú…" — sai mà vẫn chạy trơn. Khai `"dong_header": 2` để chỉ dò trên đúng dòng tiêu đề thật.
    """
    hmap, theo_dong = hmaps
    if cfg.get("dong_header"):
        hmap = theo_dong.get(cfg["dong_header"], {})
    ten = cfg.get("header")
    ung_vien = ten if isinstance(ten, list) else ([ten] if ten else [])
    for t in ung_vien:
        j = hmap.get(_nd(t))
        if j is not None:
            return j
    if cfg.get("cot") or cfg.get("cot_du_phong"):
        chu = cfg.get("cot") or cfg["cot_du_phong"]
        if cfg.get("cot_du_phong") and ung_vien:
            warn.append(f"cột '{ung_vien[0]}' không thấy theo tên -> dùng phao {chu}")
        return column_index_from_string(chu) - 1
    if cfg.get("bat_buoc", True):
        warn.append(f"THIẾU cột bắt buộc '{ten}' ({nhan})")
    return None


def _thang_cuoi(v, nam):
    """Ô chứa SỐ THÁNG (1..12) -> 'YYYY-MM-DD' ngày cuối tháng đó, năm lấy từ tên file.

    Dùng cho báo cáo xếp 12 DÒNG THÁNG (QTVH An Taxi: một file cấp số cho cả năm nên kỳ KHÔNG
    suy được từ tên file). Neo CUỐI THÁNG để `period_month` rơi đúng tháng — cùng quy ước với
    `moi_cot_thang` của bản kế hoạch. Ô là chữ ("Quý", "06 Tháng đầu năm") -> None, nhờ vậy các
    dòng tổng quý/nửa năm ngay dưới dải 12 tháng tự bị bộ lọc `ngay khác_rong` loại đi thay vì
    cộng đôi vào một kỳ nào đó.
    """
    if nam is None:
        return None
    m = re.match(r"^\s*(\d{1,2})(?:\.0)?\s*$", str(v if v is not None else "").strip())
    if not m or not 1 <= int(m.group(1)) <= 12:
        return None
    mo = int(m.group(1))
    return dt.date(nam, mo, calendar.monthrange(nam, mo)[1]).isoformat()


def _ngay_trong_thang(v, ky):
    """Ô chứa SỐ NGÀY (1..31) -> 'YYYY-MM-DD', (năm, tháng) lấy từ TÊN FILE (`_ky_thang`).

    Bản THEO DÒNG của `moi_cot_ngay` (báo cáo ngày An Taxi xếp mỗi ngày MỘT DÒNG, cột A là số
    ngày). Ngày 30/31 ở tháng ngắn -> None (dòng thừa của mẫu in sẵn), không phải lỗi.
    """
    if not ky:
        return None
    m = re.match(r"^\s*(\d{1,2})(?:\.0)?\s*$", str(v if v is not None else "").strip())
    if not m:
        return None
    try:
        return dt.date(ky[0], ky[1], int(m.group(1))).isoformat()
    except ValueError:
        return None


def _lay_o(row, j, cfg, dem_loi=None):
    v = row[j] if j < len(row) else None
    if _o_loi(v):
        if dem_loi is not None:
            k = str(v).strip().upper()
            dem_loi[k] = dem_loi.get(k, 0) + 1
        v = None
    kieu = cfg.get("kieu", "text")
    if kieu == "so":
        return _so(v, float(cfg.get("he_so", 1.0)))
    if kieu == "so_co_rong":
        return _so_co_rong(v, float(cfg.get("he_so", 1.0)))
    if kieu == "date":
        return _date(v)
    if kieu == "thang_cuoi":
        return _thang_cuoi(v, cfg.get("_nam"))
    if kieu == "ngay_trong_thang":
        return _ngay_trong_thang(v, cfg.get("_ky"))
    s = str(v).strip() if v is not None else None
    return s or None


def _dat(rec, dich, val):
    if dich.startswith("payload."):
        rec.setdefault("payload", {})[dich[len("payload."):]] = val
    else:
        rec[dich] = val


def _lay(rec, dich):
    if dich.startswith("payload."):
        return (rec.get("payload") or {}).get(dich[len("payload."):])
    return rec.get(dich)


def _qua_loc(rec, loc):
    for f in loc or []:
        v = _lay(rec, f["cot"])
        dk, gt = f.get("dieu_kien", "khac_rong"), f.get("gia_tri")
        if dk == "khac_rong" and (v is None or v == "" or v == 0):
            return False
        if dk == "bang" and str(v) != str(gt):
            return False
        if dk == "khac" and str(v) == str(gt):
            return False
        if dk == "thuoc" and str(v) not in [str(x) for x in gt]:
            return False
        if dk == "khong_thuoc" and str(v) in [str(x) for x in gt]:
            return False
        if dk == "lon_hon" and not (isinstance(v, (int, float)) and v > gt):
            return False
        if dk == "nho_hon" and not (isinstance(v, (int, float)) and v < gt):
            return False
        if dk == "chua" and _nd(gt) not in _nd(v):
            return False
    return True


def _qua_loc_o(row, dieu_kien):
    """Bản `_qua_loc` chạy THẲNG trên ô của dòng nguồn (địa chỉ theo CHỮ CỘT), dùng cho
    `ngu_canh_dong` — lúc đó bản ghi chưa được dựng nên chưa có trường nào để so.

    `khong_regex` là điều kiện chủ lực: dòng tiêu đề đơn vị nhận biết bằng "cột A có chữ nhưng
    KHÔNG phải mã chỉ tiêu" (mã dạng B110/B120…). Không có nó thì các dòng chi tiết không mã
    (Doanh thu công gò/công sơn) cũng bị nhận nhầm là tiêu đề đơn vị.
    """
    if not dieu_kien:
        return False
    for f in dieu_kien:
        j = column_index_from_string(f["cot"]) - 1
        v = row[j] if j < len(row) else None
        s = str(v).strip() if v is not None else ""
        dk, gt = f.get("dieu_kien", "khac_rong"), f.get("gia_tri")
        if dk == "khac_rong" and not s:
            return False
        if dk == "rong" and s:
            return False
        if dk == "bang" and s != str(gt):
            return False
        if dk == "khac" and s == str(gt):
            return False
        if dk == "regex" and not re.search(gt, s):
            return False
        if dk == "khong_regex" and re.search(gt, s):
            return False
        if dk == "chua" and _nd(gt) not in _nd(s):
            return False
    return True


def _dan_xuat(rec, cong_thuc):
    """eval trong sandbox hẹp — spec là cấu hình nội bộ tin cậy (xem docstring đầu file)."""
    if not cong_thuc:
        return
    env = {k: v for k, v in rec.items() if isinstance(v, (int, float))}
    env.update({f"payload.{k}": v for k, v in (rec.get("payload") or {}).items()})
    safe = {re.sub(r"\W", "_", k): (v or 0) for k, v in env.items()}
    for dich, bt in cong_thuc.items():
        try:
            _dat(rec, dich, eval(re.sub(r"\W", "_", bt) if "." in bt else bt,  # noqa: S307
                                {"__builtins__": {}}, safe))
        except Exception as ex:                                    # noqa: BLE001
            _dat(rec, dich, None)
            rec.setdefault("_loi", []).append(f"{dich}: {ex}")


def _tuan_truoc(bc):
    """Ngày phát hành báo cáo -> (thứ Hai, Chủ nhật) của TUẦN LIỀN TRƯỚC.

    Không suy từ số tuần trong tên file (8.1 / 8.2 — "tuần thứ mấy của tháng" mỗi nơi đếm một
    kiểu) mà suy từ chính ngày phát hành: lùi về thứ Hai của tuần chứa nó rồi lùi tiếp 7 ngày.
    Báo cáo ra đúng thứ Hai 03/08 -> tuần 27/07..02/08; ra muộn tới thứ Tư 05/08 vẫn ra đúng
    tuần đó, nên file về trễ không nhảy kỳ.
    """
    thu_hai = bc - dt.timedelta(days=bc.weekday())
    return thu_hai - dt.timedelta(days=7), thu_hai - dt.timedelta(days=1)


def chu_ky_tuan(spec, path):
    """-> {'ngay_bao_cao','tuan_tu','tuan_den'} để nhét vào payload, hoặc {} nếu spec không khai.

    Số trong file là DƯ NỢ TẠI THỜI ĐIỂM chốt tuần, không phải phát sinh trong tuần. Giao diện
    phải nói được "ảnh chụp tuần nào, phát hành ngày nào" — thiếu ba mốc này thì người xem không
    phân biệt được số của tuần trước với số mới nhất.
    """
    c = spec.get("ky_tu_ten_file") or {}
    if not (c.get("regex_ngay") and c.get("tuan_truoc_ngay_bao_cao")):
        return {}
    ten = os.path.basename(path)
    mn, mt = re.search(c["regex_nam"], ten), re.search(c["regex_thang"], ten)
    md = re.search(c["regex_ngay"], ten)
    if not (mn and mt and md):
        return {}
    try:
        bc = dt.date(int(mn.group(1)), int(mt.group(1)), int(md.group(1)))
    except ValueError:
        return {}
    tu, den = _tuan_truoc(bc)
    return {"ngay_bao_cao": bc.isoformat(), "tuan_tu": tu.isoformat(), "tuan_den": den.isoformat()}


def ngay_tu_ten_file(spec, path):
    """-> ('YYYY-MM-DD' | None, [cảnh báo]). Dùng cho bảng ẢNH CHỤP (không có cột ngày từng dòng).

    `ky_tu_ten_file` tách RIÊNG regex năm và tháng — file công nợ đặt tên
    "…M.2026.07.22_Baocaocongnophaithu_T1.xlsx": cụm 2026.07.22 là NGÀY LẬP báo cáo (giống hệt
    nhau ở mọi kỳ), kỳ thật nằm ở hậu tố _T1/_T7. Lấy nhầm cụm đầu là dồn cả 12 tháng vào tháng 7.
    `ngay_tu_ten_file` dùng khi tên file có ngày chốt thật (kho xe B2B: "…2026.7.31.KHO XE.xlsx").

    `regex_ngay` (tuỳ chọn, thêm 12/08/2026): tên file báo cáo TUẦN của KSCL ghi đủ
    "W.<năm>.<tháng>.<tuần>.<ngày báo cáo>" — vd "W.2026.8.2.10" = tuần 2 tháng 8, phát hành
    ngày 10/08. Trước đây chỉ đọc năm+tháng rồi gán CUỐI THÁNG, nên mọi bản tuần của một tháng
    đổ chung vào 31/08: lọc theo tuần không ra gì, và nhiều tuần thì phải vứt bớt file. Có
    `regex_ngay` thì kỳ lấy đúng theo ngày (xem thêm `tuan_truoc_ngay_bao_cao`).
    """
    ten = os.path.basename(path)
    warn = []
    if spec.get("ky_tu_ten_file"):
        c = spec["ky_tu_ten_file"]
        mn, mt = re.search(c["regex_nam"], ten), re.search(c["regex_thang"], ten)
        md = re.search(c["regex_ngay"], ten) if c.get("regex_ngay") else None
        if mn and mt and c.get("regex_ngay") and not md:
            warn.append(f"không dò được NGÀY báo cáo từ tên file: {ten} — tạm lấy cuối tháng")
        if mn and mt and md:
            try:
                bc = dt.date(int(mn.group(1)), int(mt.group(1)), int(md.group(1)))
            except ValueError:
                warn.append(f"ngày báo cáo trong tên file không hợp lệ: {md.group(0)}")
            else:
                if not c.get("tuan_truoc_ngay_bao_cao"):
                    return bc.isoformat(), warn
                tu, den = _tuan_truoc(bc)
                if bc.weekday() != 0:
                    # Quy ước: báo cáo phát hành THỨ HAI, chốt số của tuần liền trước. File ra
                    # muộn 1-2 hôm vẫn quy về đúng tuần đó (lùi về thứ Hai của tuần chứa nó), còn
                    # ra vào cuối tuần thì rất có thể là chốt CHÍNH tuần đang chạy -> phải kêu,
                    # đừng lặng lẽ gán nhầm một tuần.
                    warn.append(f"ngày báo cáo {bc.isoformat()} KHÔNG phải thứ Hai "
                                f"({bc.strftime('%A')}) — vẫn quy về tuần {tu}..{den}, kiểm tra lại")
                return den.isoformat(), warn
        if mn and mt:
            y, mo = int(mn.group(1)), int(mt.group(1))
            return dt.date(y, mo, calendar.monthrange(y, mo)[1]).isoformat(), warn
        warn.append(f"không dò được kỳ (năm/tháng) từ tên file: {ten}")
    if spec.get("ngay_tu_ten_file"):
        m = re.search(spec["ngay_tu_ten_file"]["regex"], ten)
        if m:
            try:
                return dt.date(*(int(m.group(i + 1)) for i in range(3))).isoformat(), warn
            except ValueError:
                warn.append(f"ngày trong tên file không hợp lệ: {m.group(0)}")
        else:
            warn.append(f"không dò được ngày từ tên file: {ten}")
    return None, warn


def _ky_thang(spec, path):
    """-> ((năm, tháng), [cảnh báo]) cho `ban_ghi = moi_cot_ngay`. Kỳ lấy từ TÊN FILE
    (vd '...202608.Baocaodoanhthungay.xlsx'), vì các cột chỉ ghi số ngày, không ghi tháng."""
    c = spec.get("ky_thang_tu_ten_file") or {"regex": r"\.(\d{4})(\d{2})\."}
    m = re.search(c["regex"], os.path.basename(path))
    if not m:
        return None, [f"không dò được kỳ (năm/tháng) từ tên file: {os.path.basename(path)}"]
    return (int(m.group(1)), int(m.group(2))), []


def _chuyen_xls_cu(duong_dan):
    """`.xls` đời cũ (BIFF) -> sinh bản `.xlsx` cạnh nó, trả đường dẫn mới (None nếu hỏng).

    openpyxl KHÔNG đọc được định dạng Excel 97-2003; nghiệp vụ thỉnh thoảng lưu nhầm (công nợ
    phải thu T5 gửi 12/08/2026). Không có bước này thì file nằm im, không dòng nào vào DB và
    KHÔNG có cảnh báo nào — đúng kiểu thiếu số mà không ai biết.

    Bản chuyển đổi lưu CẠNH file gốc (cùng thư mục -> `source_file` trong DB vẫn là
    "<FOLDER>::<tên>.xlsx" đúng quy ước) và dùng lại ở lần quét sau, chỉ chuyển khi thiếu hoặc
    cũ hơn bản .xls.

    Dùng xlrd + openpyxl chứ KHÔNG gọi `soffice`: máy này chỉ cài libreoffice-core/common, thiếu
    hẳn gói `libreoffice-calc` nên soffice trả "source file could not be loaded" với MỌI bảng
    tính, kể cả .xlsx hợp lệ (đã thử 12/08/2026). Đường thuần Python còn khỏi cần sudo để cài gói
    và khỏi đẻ tiến trình con giữa lượt nạp.

    Chỉ bê GIÁ TRỊ ô sang, không giữ định dạng — engine chỉ đọc giá trị (`data_only=True`).
    """
    dich = os.path.splitext(duong_dan)[0] + ".xlsx"
    if os.path.exists(dich) and os.path.getmtime(dich) >= os.path.getmtime(duong_dan):
        return dich
    try:
        import xlrd
    except ImportError:
        return None
    try:
        nguon = xlrd.open_workbook(duong_dan)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sh in nguon.sheets():
            # Tên sheet: Excel giới hạn 31 ký tự và cấm : \ / ? * [ ]
            ws = wb.create_sheet(title=re.sub(r"[:\\/?*\[\]]", "-", sh.name)[:31] or "Sheet")
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    o = sh.cell(r, c)
                    if o.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                        continue
                    v = o.value
                    if o.ctype == xlrd.XL_CELL_DATE:
                        try:
                            v = dt.datetime(*xlrd.xldate_as_tuple(v, nguon.datemode))
                        except Exception:                              # noqa: BLE001
                            pass                                       # số ngày hỏng -> để nguyên
                    elif o.ctype == xlrd.XL_CELL_BOOLEAN:
                        v = bool(v)
                    ws.cell(row=r + 1, column=c + 1, value=v)
        wb.save(dich)
    except Exception:                                                  # noqa: BLE001
        return None
    return dich


def quet_nguon(spec):
    """Danh sách file của nguồn -> [đường dẫn], [cảnh báo].

    KHỚP ĐUÔI KHÔNG PHÂN BIỆT HOA/THƯỜNG. `glob` phân biệt hoa thường trên Linux, nên
    `"*.xlsx"` bỏ qua sạch `.Xlsx` — nghiệp vụ đặt tên lẫn lộn cả hai kiểu. Đã mất 4/7 kỳ công
    nợ phải thu SRVF (T2/T3/T4/T6 đều `.Xlsx`) suốt từ lúc dựng nguồn tới 12/08/2026, không một
    dòng cảnh báo: file có trên đĩa, spec chạy "thành công", màn hình chỉ trống mấy tháng giữa.
    """
    warn = []
    thu_muc = os.path.join(REPORTS_DIR, (spec.get("nguon") or {}).get("folder", ""))
    if not os.path.isdir(thu_muc):
        return [], [f"không có thư mục nguồn: {thu_muc}"]
    mau = ((spec.get("nguon") or {}).get("file_glob") or "*.xlsx").lower()
    ten = sorted(os.listdir(thu_muc))
    # File tạm Excel sinh ra khi ai đó đang MỞ file trên máy chia sẻ. Khớp "*.xlsx" nhưng không
    # phải workbook thật -> đọc vào là ném lỗi khó hiểu giữa lượt nạp.
    ten = [n for n in ten if not n.startswith("~$")]
    out = [os.path.join(thu_muc, n) for n in ten if fnmatch.fnmatch(n.lower(), mau)]
    if mau.endswith(".xlsx"):
        da_co = {os.path.splitext(p)[0].lower() for p in out}
        for n in ten:
            g = os.path.join(thu_muc, n)
            if n.lower().endswith(".xls") and os.path.splitext(g)[0].lower() not in da_co:
                moi = _chuyen_xls_cu(g)
                if moi:
                    warn.append(f"{n}: Excel 97-2003, đã chuyển sang {os.path.basename(moi)}")
                    out.append(moi)
                else:
                    warn.append(f"{n}: Excel 97-2003 và CHUYỂN ĐỔI HỎNG — file này chưa vào DB")
    return sorted(out), warn


def loc_file_moi_nhat(spec, files):
    """Mỗi KỲ chỉ giữ file có ngày chốt MỚI NHẤT.

    Kho xe B2B có 2 ảnh chụp cùng tháng 7 ("…7.24.KHO XE" và "…7.31.KHO XE"). Nạp cả hai là
    ĐẾM ĐÔI tồn kho tháng 7 (mỗi VIN xuất hiện 2 lần) — số tồn phình gần gấp đôi mà không có
    cảnh báo nào. Bật `"moi_ky_lay_file_moi_nhat": true` trong spec cho mọi nguồn ảnh chụp.

    `"moi_ky_lay_file_moi_nhat": "ngay"` — gộp theo NGÀY CHỐT thay vì theo tháng, cho nguồn báo
    cáo TUẦN (KSCL): các bản tuần trong cùng một tháng là những ảnh chụp KHÁC NHAU, phải giữ đủ
    để lọc theo tuần ra được số của đúng tuần đó. Gộp theo tháng ở đây là vứt mất tuần cũ. Việc
    không cộng đôi khi xem cả tháng do tầng đọc lo (`_SNAP_RT` -> chỉ lấy ngày chốt mới nhất).

    `"moi_ky_lay_file_moi_nhat": "mot_file"` (14/08/2026) — giữ ĐÚNG MỘT file mới nhất trong CẢ
    thư mục, bất kể kỳ. Cho nguồn mà MỖI FILE THÁNG ĐÃ CHỨA TOÀN BỘ CÁC KỲ TỪ ĐẦU NĂM: sheet
    "Chi tiết" của báo cáo nhân sự HCNS có đủ 6 tháng trong file tháng 1 lẫn file tháng 6. Hai chế
    độ trên không cứu được vì mỗi file suy ra một kỳ KHÁC NHAU nên đều được giữ -> nạp 6 file là
    ghi cùng dữ liệu 6 lần dưới 6 `source_file`, và `_ghi` (xoá theo source_file) không chặn nổi:
    tổng nhân sự phình 6 lần.
    """
    che_do = spec.get("moi_ky_lay_file_moi_nhat")
    if not che_do:
        return files, []
    giu, bo = {}, []
    for f in files:
        ngay, _ = ngay_tu_ten_file(spec, f)
        if not ngay:
            giu[f] = f            # không suy được kỳ -> giữ nguyên, đừng im lặng loại
            continue
        ky = "" if che_do == "mot_file" else (ngay if che_do == "ngay" else ngay[:7])
        cu = giu.get(ky)
        # HOÀ ngày chốt -> lấy file VỀ SAU (mtime). Nghiệp vụ gửi lại bản SỬA của cùng một kỳ với
        # tên khác (claim T3-T6: "…7.24. BaocaoClaim_B2C_T3" rồi "…8.11. BaocaoClaim_B2C_T3"),
        # hai tên cùng suy ra một kỳ nên so ngày là hoà. Không phá hoà thì thứ tự sorted() quyết
        # định, mà "7.24" đứng trước "8.11" -> giữ đúng bản CŨ và vứt bản đã sửa.
        if cu is None:
            moi_hon = True
        else:
            ngay_cu = ngay_tu_ten_file(spec, cu)[0]
            moi_hon = (ngay, os.path.getmtime(f)) > (ngay_cu, os.path.getmtime(cu))
        if moi_hon:
            if cu is not None:
                bo.append(os.path.basename(cu))
            giu[ky] = f
        else:
            bo.append(os.path.basename(f))
    return sorted(giu.values()), bo


# ─────────────────────────── trích 1 file ───────────────────────────
def _tron_vung(spec, v):
    """Trộn 1 khai báo `vung` lên spec gốc -> spec con chỉ đọc ĐÚNG khối đó.

    Khoá trong `vung` GHI ĐÈ khoá cùng tên ở spec, riêng `cot` / `chieu_co_dinh` thì trộn theo
    từng khoá con (vùng chỉ cần khai lại cột nào LỆCH, không phải chép cả bảng cột).
    """
    sp = {k: x for k, x in spec.items() if k != "vung"}
    for k, x in v.items():
        if k in ("cot", "chieu_co_dinh"):
            sp[k] = {**(spec.get(k) or {}), **(x or {})}
        elif k != "ten":
            sp[k] = x
    return sp


def extract_file(spec, path):
    """Trích 1 file. `vung` (nếu có) = NHIỀU KHỐI trong CÙNG một sheet, đọc lần lượt.

    VÌ SAO CÓ `vung` (14/08/2026, báo cáo QTVH An Taxi): file `B.7.AAG.PKDVH.M.2026...` xếp 6 khối
    nghiệp vụ (A nhân sự · B đội xe · C chuyến · D1 doanh thu · D2 thuê xe) NỐI TIẾP nhau trong
    một sheet, mỗi khối có dòng tiêu đề riêng và 12 dòng tháng, lại chia tiếp theo depot
    (I Sơn Tây · II Thái Nguyên · III Tổng 2 depot). Một spec = một dải dòng nên nếu không có
    `vung` thì phải đẻ ~10 file spec gần trùng nhau cho CÙNG một file nguồn — mỗi lần mapping đổi
    là sửa 10 chỗ. Nay 1 spec khai `cot` chung + liệt kê các vùng lệch.

    Mỗi vùng vẫn tự dò cột THEO TÊN HEADER trong đúng dòng tiêu đề của nó (nguyên tắc số 1) và
    vẫn kiểm được `kiem_tra_o` riêng, nên chèn/xoá cột trong một khối không lây sang khối khác.
    """
    vung = spec.get("vung")
    if not vung:
        return _extract_vung(spec, path)
    recs, warn = [], []
    for i, v in enumerate(vung, 1):
        r, w = _extract_vung(_tron_vung(spec, v), path)
        recs += r
        warn += [f"[vùng {v.get('ten') or i}] {x}" for x in w]
    return recs, warn


def _extract_vung(spec, path):
    warn, recs = [], []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        _sh_cfg = (spec.get("nguon") or {}).get("sheet") or {}
        _thang = None
        if "theo_thang" in _sh_cfg:
            _ky, _w = _ky_thang(spec, path)
            warn += _w
            _thang = _ky[1] if _ky else None
        sheet = _chon_sheet(wb, _sh_cfg, thang=_thang)
        if not sheet:
            return [], [f"không chọn được sheet (có: {wb.sheetnames})"]
        ws = wb[sheet]
        hdr_cfg = spec.get("header") or {"dong": 1}
        if hdr_cfg.get("khong_co"):
            # Sheet KHÔNG có dòng tiêu đề (vd "Tồn kho xe vật lý": dữ liệu chạy thẳng từ dòng 2,
            # 3 cột không tên). Khi đó mọi khai báo cột PHẢI dùng `"cot": "<chữ cột>"`.
            max_hdr, hmap = 0, ({}, {})
        elif hdr_cfg.get("tim_o"):
            # TỰ DÒ dòng header theo một nhãn mốc. Cần khi cùng một spec phải đọc nhiều sheet mà
            # dòng header nằm ở vị trí KHÁC NHAU: BCTC SRVF để header ở dòng 1 tại sheet T01 nhưng
            # dòng 8 tại T07 (T07 có thêm 7 dòng tiêu đề đơn vị/địa chỉ phía trên). Khai `dong`
            # cứng thì một trong hai sheet chắc chắn trượt; gộp `dong: [1,8]` cũng sai vì ô đầu
            # dòng 1 của T07 là tên chi nhánh, không phải "Mã số".
            moc = _nd(hdr_cfg["tim_o"])
            toi_da = int(hdr_cfg.get("toi_da", 30))
            quet = [list(r) for r in ws.iter_rows(min_row=1, max_row=toi_da, values_only=True)]
            max_hdr = 0
            for i, r in enumerate(quet, start=1):
                if any(_nd(c) == moc for c in r if c not in (None, "")):
                    max_hdr = i
                    break
            if not max_hdr:
                return [], [f"không tìm được dòng header chứa '{hdr_cfg['tim_o']}' "
                            f"trong {toi_da} dòng đầu sheet '{sheet}'"]
            hmap = _map_header(quet[:max_hdr], max_hdr)
        else:
            hdr_dong = hdr_cfg.get("dong", 1)
            max_hdr = max(hdr_dong) if isinstance(hdr_dong, list) else hdr_dong
            head_rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=max_hdr,
                                                       values_only=True)]
            hmap = _map_header(head_rows, hdr_dong)
        cot = _resolve_cot(spec, hmap, warn)

        # Kỳ nằm ở TÊN FILE nhưng số tháng/ngày nằm TRONG Ô -> ghép hai nửa lại ngay tại đây rồi
        # nhét vào cfg của cột, để `_lay_o` (chỉ thấy 1 ô) dựng được ngày đầy đủ. Bản kế hoạch báo
        # cáo QTVH An Taxi là MỘT file cho cả năm ("…M.2026.Baocaotonghop"), nên kỳ tuyệt đối
        # không được suy từ tên file như các nguồn ảnh-chụp khác.
        _nam, _ky_ngay = None, None
        if any(c.get("kieu") == "thang_cuoi" for _, c in cot.values()):
            cn = spec.get("nam_tu_ten_file") or {"regex": r"\.M\.(\d{4})\."}
            mn = re.search(cn["regex"], os.path.basename(path))
            if mn:
                _nam = int(mn.group(1))
            else:
                warn.append(f"không dò được NĂM từ tên file: {os.path.basename(path)}")
        if any(c.get("kieu") == "ngay_trong_thang" for _, c in cot.values()):
            _ky_ngay, w5 = _ky_thang(spec, path)
            warn.extend(w5)
        for dich, (j, c) in list(cot.items()):
            if c.get("kieu") == "thang_cuoi":
                cot[dich] = (j, {**c, "_nam": _nam})
            elif c.get("kieu") == "ngay_trong_thang":
                cot[dich] = (j, {**c, "_ky": _ky_ngay})

        chieu = {}
        for dich, cfg in (spec.get("chieu_tu_ten_file") or {}).items():
            m = re.search(cfg["regex"], os.path.basename(path), re.I)
            v = (m.group(int(cfg.get("nhom", 1))) if m else cfg.get("mac_dinh"))
            chieu[dich] = (str(v).upper() if v and cfg.get("hoa") else v)

        # CHỐT CHẶN LAYOUT: nguồn chỉ đúng khi các ô mốc khớp. Bắt buộc với spec dùng CHỮ CỘT —
        # file công nợ T1 và T7 tuy cùng tên sheet nhưng bố cục KHÁC HẲN (T1 không có phân tách
        # kênh, header ở dòng 5). Không có chốt này thì T1 vẫn "chạy được" và đẻ ra 9,86 tỷ vô
        # nghĩa, im lặng — kiểu sai nguy hiểm nhất.
        for ktr in spec.get("kiem_tra_o") or []:
            thuc = ws[ktr["o"]].value
            if _nd(ktr.get("bang")) not in _nd(thuc):
                return [], [f"BỎ QUA — layout khác spec: ô {ktr['o']} = "
                            f"{str(thuc)[:40]!r}, cần chứa {ktr.get('bang')!r}"]

        ngay_file, w2 = ngay_tu_ten_file(spec, path)
        warn.extend(w2)
        # Nguồn báo cáo TUẦN: kèm mốc tuần vào payload từng dòng để giao diện nói rõ đang xem
        # ảnh chụp tuần nào (`ngay` chỉ mang ngày CHỐT — Chủ nhật cuối tuần được phủ).
        payload_tuan = chu_ky_tuan(spec, path)

        bat_dau = spec.get("dong_bat_dau") or (max_hdr + 1)
        # `dong_ket_thuc` (14/08/2026): CHẶN TRÊN của dải dòng — bắt buộc khi nhiều khối nằm nối
        # tiếp trong một sheet (xem `vung`). Không có nó thì khối A quét luôn xuống khối B: cột
        # "Tổng" của bảng nhân sự và cột "Tổng số xe" của bảng đội xe cùng nằm ở E/C nên số của
        # khối sau vẫn "đọc được" và cộng vào khối trước — sai mà không một dấu hiệu nào.
        ket_thuc = spec.get("dong_ket_thuc")
        gia_tri_cols = spec.get("cot_gia_tri") or []
        gt_idx = [(_tim_cot(hmap, c, f"cột giá trị {c.get('dim1') or c.get('header')}", warn), c)
                  for c in gia_tri_cols]
        gt_idx = [(j, c) for j, c in gt_idx if j is not None]

        # Cột-theo-ngày: đọc SỐ NGÀY từ chính dòng tiêu đề của từng cột rồi ghép với kỳ của file.
        # KHÔNG đánh số ngày theo thứ tự cột: tháng 2 chỉ có 28-29 cột có nghĩa, và vài file chèn
        # thêm cột phụ giữa dải ngày — suy theo vị trí sẽ lệch ngày mà không có dấu hiệu gì.
        # 3 khoá TUỲ CHỌN thêm 10/08/2026 cho bản KẾ HOẠCH doanh thu theo ngày của XDV
        # (Baocaodoanhthukehoachngay, sheet "KHDT theo ngày") — mặc định giữ nguyên hành vi cũ
        # nên mọi spec đang chạy không đổi một ly:
        #   "dong": 9       -> đọc mốc ngày từ DÒNG KHÁC dòng tiêu đề. File đó ghi ngày ở dòng 9
        #                      (ô gộp 2 cột), còn dòng 10 là tiêu đề con "Số lượng"/"Doanh thu".
        #   ô là NGÀY THẬT  -> chấp nhận cell datetime/date, không chỉ số "1".."31". Kèm kiểm
        #                      tháng/năm khớp kỳ của file: lệch thì BỎ cột đó và cảnh báo, vì
        #                      nhận sai là gán số của tháng khác vào kỳ này mà không ai thấy.
        #   "lech_amount"   -> giá trị KHÔNG nằm ở cột mang mốc ngày mà lệch sang phải n cột.
        #                      Ở file này mỗi ngày chiếm 2 cột: (Số lượng, Doanh thu).
        #   "lech_amount2"  -> cột thứ hai của cặp, ghi vào amount2 (giữ luôn KH số lượng RO để
        #                      tính "doanh thu bình quân/RO mục tiêu" mà không cần nạp 2 lần).
        nc_ngay = spec.get("cot_ngay") or {}
        ngay_theo_cot = []
        if spec.get("ban_ghi") == "moi_cot_ngay":
            nam_thang, w3 = _ky_thang(spec, path)
            warn.extend(w3)
            if nam_thang:
                y, mo = nam_thang
                j1 = column_index_from_string(nc_ngay.get("tu", "G")) - 1
                j2 = column_index_from_string(nc_ngay.get("den", "AK")) - 1
                dong_ngay = nc_ngay.get("dong")
                if dong_ngay:
                    moc_row = [c.value for c in ws[int(dong_ngay)]]
                else:
                    moc_row = head_rows[(hdr_dong[-1] if isinstance(hdr_dong, list) else hdr_dong) - 1]
                for j in range(j1, j2 + 1):
                    raw = moc_row[j] if j < len(moc_row) else None
                    ngay_cot = None
                    if isinstance(raw, (dt.datetime, dt.date)):
                        if (raw.year, raw.month) == (y, mo):
                            ngay_cot = dt.date(raw.year, raw.month, raw.day).isoformat()
                        else:
                            warn.append(f"cột {get_column_letter(j + 1)} ghi ngày {raw} không thuộc "
                                        f"kỳ {y}-{mo:02d} của file -> bỏ cột")
                    else:
                        m = re.match(r"^\s*(\d{1,2})\s*$", str(raw or ""))
                        if m:
                            try:
                                ngay_cot = dt.date(y, mo, int(m.group(1))).isoformat()
                            except ValueError:
                                ngay_cot = None   # ngày 30/31 ở tháng ngắn -> cột thừa, bỏ qua
                    if ngay_cot:
                        ngay_theo_cot.append((j, ngay_cot))
            if not ngay_theo_cot:
                return [], [*warn, "BỎ QUA — không dựng được dải cột theo ngày"]

        # Cột-theo-THÁNG: bảng KẾ HOẠCH xếp 12 cột "Tháng 1".."Tháng 12" — MỘT file cấp kế hoạch
        # cho cả năm, nên một lần nạp phải đẻ ra 12 kỳ. Tháng đọc từ CHÍNH tiêu đề cột (không suy
        # theo vị trí): trước dải tháng còn có "Tổng năm" / "6 tháng đầu năm" / "6 tháng cuối năm"
        # và mấy cột tổng này rất hay bị chèn thêm giữa các bản kế hoạch. Năm lấy từ tên file.
        nc_thang = spec.get("cot_thang") or {}
        thang_theo_cot = []
        if spec.get("ban_ghi") == "moi_cot_thang":
            nam_thang, w4 = _ky_thang(spec, path)
            warn.extend(w4)
            if nam_thang:
                y = nam_thang[0]
                j1 = column_index_from_string(nc_thang.get("tu", "G")) - 1
                j2 = column_index_from_string(nc_thang.get("den", "R")) - 1
                hdr_row = head_rows[(hdr_dong[-1] if isinstance(hdr_dong, list) else hdr_dong) - 1]
                # `regex` (tuỳ chọn, 13/08/2026): bản kế hoạch XDV không ghi tiêu đề cột trần là
                # "Tháng 1" mà kèm tên chỉ tiêu — "Sản lượng Tháng 1", "Doanh thu Tháng 1", "Lợi
                # nhuận gộp\nTháng 1" (mỗi mục trong sheet một kiểu). Mặc định giữ nguyên khớp
                # CHẶT `^thang\d+$` để các spec đang chạy không đổi hành vi; spec nào cần thì nới
                # thành `thang(\d{1,2})$`. Vẫn chỉ dò trong dải cột `tu`..`den`, nên "Tổng 6 tháng
                # đầu năm" (không kết thúc bằng số) và các cột tổng khác không lọt vào.
                re_thang = re.compile(nc_thang.get("regex") or r"^thang(\d{1,2})$")
                for j in range(j1, j2 + 1):
                    m = re_thang.search(_nd(hdr_row[j] if j < len(hdr_row) else None))
                    if not m or not 1 <= int(m.group(1)) <= 12:
                        continue
                    mo = int(m.group(1))
                    # Kỳ của kế hoạch là CẢ THÁNG -> neo vào ngày cuối tháng, giống mọi nguồn
                    # ảnh chụp khác, để `period_month` rơi đúng tháng đó.
                    thang_theo_cot.append(
                        (j, dt.date(y, mo, calendar.monthrange(y, mo)[1]).isoformat()))
            if not thang_theo_cot:
                return [], [*warn, "BỎ QUA — không dựng được dải cột theo tháng"]

        khong_map, bo_loc, o_loi = {}, 0, {}
        ngu_canh = {}          # ngữ cảnh mang từ dòng tiêu đề xuống, xem `ngu_canh_dong`
        nc_cfg = spec.get("ngu_canh_dong")
        for row in ws.iter_rows(min_row=bat_dau, max_row=ket_thuc, values_only=True):
            # BẢNG PHÂN CẤP: một số báo cáo không lặp lại tên đơn vị trên từng dòng mà đặt nó ở
            # DÒNG TIÊU ĐỀ riêng, các dòng bên dưới ngầm hiểu là của đơn vị đó (báo cáo doanh thu
            # XDV: dòng "3S có đồng sơn | Ocean Park" rồi 8 dòng mã B110..B150 bên dưới).
            # Dòng tiêu đề KHÔNG tự sinh bản ghi — nó chỉ đặt ngữ cảnh.
            # `ngu_canh_dong` nhận 1 quy tắc (dict) hoặc NHIỀU quy tắc (list) — file doanh thu XDV
            # có 2 TẦNG ngữ cảnh lồng nhau: dòng "II/ DOANH THU LỆNH W (BẢO HÀNH)" mở một khối
            # loại lệnh, bên trong lại có các dòng tên xưởng. Quy tắc xét theo thứ tự, khớp cái
            # đầu tiên; `xoa` liệt kê các trường phải quên khi sang khối mới (đổi khối loại lệnh
            # thì cost_center của xưởng cuối khối trước KHÔNG được mang sang).
            hit = False
            for rule in ([nc_cfg] if isinstance(nc_cfg, dict) else (nc_cfg or [])):
                if not _qua_loc_o(row, rule.get("khi") or []):
                    continue
                for k in rule.get("xoa") or []:
                    ngu_canh.pop(k, None)
                for dich, c in (rule.get("gan") or {}).items():
                    v = _lay_o(row, column_index_from_string(c["cot"]) - 1, c, o_loi)
                    if c.get("anh_xa"):
                        # Nhãn trong file dài dòng ("II/ DOANH THU LỆNH W (BẢO HÀNH)") -> quy về
                        # giá trị ngắn dùng cho dim. So theo kiểu "chứa", bỏ dấu.
                        v = next((out for key, out in c["anh_xa"].items() if _nd(key) in _nd(v)),
                                 c.get("mac_dinh"))
                    hook = c.get("chuan_hoa")
                    if hook and v is not None:
                        res = _CHUAN_HOA[hook](v)
                        if isinstance(res, dict):
                            if res.get("_khong_map"):
                                khong_map[res["_khong_map"]] = khong_map.get(res["_khong_map"], 0) + 1
                            ngu_canh.update({k: x for k, x in res.items() if k != "_khong_map"})
                            continue
                        v = res
                    ngu_canh[dich] = v
                hit = True
                break
            if hit:
                continue          # dòng tiêu đề KHÔNG tự sinh bản ghi
            base = {"payload": {**(spec.get("payload_them") or {}), **payload_tuan}}
            if spec.get("khoi"):
                base["khoi"] = spec["khoi"]
            base.update(spec.get("chieu_co_dinh") or {})   # vd {"dim2": "B2B"} cho file 1 kênh
            for dich, v in ngu_canh.items():
                _dat(base, dich, v)
            base.update({k: v for k, v in chieu.items()})
            if ngay_file:
                base["ngay"] = ngay_file
            for dich, (j, cfg) in cot.items():
                val = _lay_o(row, j, cfg, o_loi)
                hook = cfg.get("chuan_hoa")
                if hook and val is not None:
                    res = _CHUAN_HOA[hook](val)
                    if isinstance(res, dict):
                        # Hook hỏng KHÔNG loại dòng ngay: phải để `loc` chạy trước, nếu không mọi
                        # dòng RÁC (vùng dán thừa, không có ngày) sẽ bị gộp vào cảnh báo "không map
                        # được đơn vị" và che mất đơn vị hỏng THẬT. Đánh dấu rồi xử lý sau bộ lọc.
                        if res.get("_khong_map"):
                            # `giu_khi_khong_map`: dòng KHÔNG khớp danh mục vẫn là dòng THẬT và
                            # phải giữ để tổng không hụt — vd 16 xe tồn đứng tên "CHI NHÁNH
                            # VINFAST HÀ NỘI" trong báo cáo tồn vật lý: không thuộc SR nào nhưng
                            # là xe có thật. Giữ lại, cost_center để trống, tên thô nằm ở payload.
                            if not cfg.get("giu_khi_khong_map"):
                                base["_khong_map"] = res["_khong_map"]
                        for k2, v2 in res.items():
                            if k2 != "_khong_map":
                                base[k2] = v2
                        continue
                    val = res
                _dat(base, dich, val)
            if all(base.get(k) in (None, "") for k in _CHUAN_TRUONG if k != "khoi") \
                    and not (base.get("payload") or {}):
                continue
            # DỪNG QUÉT khi gặp dòng thoả `dung_khi` — sheet "CN theo đơn vị" có dòng "Total" rồi
            # BÊN DƯỚI còn một bảng phụ (công nợ quá hạn đã có COC) cũng liệt kê tên Showroom ở
            # cột A. Không dừng là gộp luôn bảng phụ vào -> cộng đôi mà không có dấu hiệu gì.
            if spec.get("dung_khi") and _qua_loc(base, spec["dung_khi"]):
                break
            outs = []
            if spec.get("ban_ghi") == "moi_cot_ngay":
                # Bảng có MỘT CỘT MỖI NGÀY (báo cáo doanh thu XDV: G..AK = ngày 01..31, tiêu đề
                # cột chính là số ngày). Mỗi ô có số -> 1 bản ghi mang đúng `ngay` của cột đó,
                # nhờ vậy vừa vẽ được biểu đồ theo ngày vừa cộng lên tháng mà không đếm đôi.
                # Ô rỗng/0 bị bỏ: tháng đang dở thì các ngày chưa tới đều là 0, giữ lại chỉ làm
                # phình bảng và đẻ ra "ngày có số 0" giả.
                lech = int(nc_ngay.get("lech_amount", 0))
                lech2 = nc_ngay.get("lech_amount2")
                for j, ngay_cot in ngay_theo_cot:
                    jv = j + lech
                    v = _so(row[jv] if jv < len(row) else None, float(nc_ngay.get("he_so", 1.0)))
                    if not v:
                        continue
                    r2 = json.loads(json.dumps(base))
                    r2["ngay"] = ngay_cot
                    r2["amount"] = v
                    if lech2 is not None:
                        j2v = j + int(lech2)
                        # he_so RIÊNG: cột thứ hai của cặp là SỐ LƯỢNG RO (đơn vị lệnh), không
                        # phải tiền — dùng chung he_so 1e-9 của tiền là ra 0.0000005 lệnh.
                        r2["amount2"] = _so(row[j2v] if j2v < len(row) else None,
                                            float(nc_ngay.get("he_so_amount2", 1.0)))
                    outs.append(r2)
            elif spec.get("ban_ghi") == "moi_cot_thang":
                # Ô rỗng/0 bị bỏ: bản kế hoạch để trống các tháng chưa đăng ký (A230/A250 trống
                # hết T1-T6). Giữ lại là đẻ ra "kế hoạch = 0" giả, và %HT sẽ chia cho 0.
                for j, ngay_cot in thang_theo_cot:
                    v = _so(row[j] if j < len(row) else None, float(nc_thang.get("he_so", 1.0)))
                    if not v:
                        continue
                    r2 = json.loads(json.dumps(base))
                    r2["ngay"] = ngay_cot
                    r2["amount"] = v
                    outs.append(r2)
            elif spec.get("ban_ghi") == "moi_cot_gia_tri":
                for j, c in gt_idx:
                    r2 = json.loads(json.dumps(base))
                    r2["amount"] = _so(row[j] if j < len(row) else None,
                                       float(c.get("he_so", 1.0)))
                    for k2 in ("dim1", "dim2", "dim3"):
                        if c.get(k2):
                            r2[k2] = c[k2]
                    outs.append(r2)
            else:
                outs.append(base)
            for r2 in outs:
                _dan_xuat(r2, spec.get("dan_xuat"))
                hong = r2.pop("_khong_map", None)
                # Đơn vị KHÔNG map được mà dòng CÓ SỐ -> luôn báo ra, KHÔNG để bộ lọc nuốt trước.
                # Bắt được 2026-08-09: spec lọc `cost_center khác rỗng`, nên xưởng "Quận 12" (có
                # thật trong file DMS, chưa có trong master_data) rơi vào nhánh "bỏ N dòng không
                # qua bộ lọc" — mất 19 lệnh mà không một cảnh báo nào. Dòng không map mà TOÀN 0
                # thì vẫn coi là rác (dòng tổng/дòng trống), đếm vào bo_loc cho đỡ nhiễu.
                if hong and (r2.get("amount") or 0):
                    khong_map[hong] = khong_map.get(hong, 0) + 1
                elif hong or not _qua_loc(r2, spec.get("loc")):
                    bo_loc += 1
                else:
                    recs.append(r2)
        if bo_loc:
            warn.append(f"bỏ {bo_loc} {_W_BO_LOC}")
        if o_loi:
            # Ô lỗi Excel đã bị coi là trống ở trên — báo ra để biết FILE NGUỒN chưa cập nhật
            # xong, đừng đi tìm lỗi ở spec/deriver.
            warn.append("Ô LỖI EXCEL trong file nguồn (đã coi như ô trống): "
                        + ", ".join(f"{k} ({v} ô)"
                                    for k, v in sorted(o_loi.items(), key=lambda x: -x[1])))
        if khong_map:
            warn.append("KHÔNG MAP ĐƯỢC đơn vị (đã qua bộ lọc, tức là dòng THẬT): "
                        + ", ".join(f"{k} ({v} dòng)"
                                    for k, v in sorted(khong_map.items(), key=lambda x: -x[1])[:10]))
        return recs, warn
    finally:
        wb.close()


def _source_id(path):
    parts = os.path.normpath(path).split(os.sep)
    return f"{parts[-3]}::{os.path.basename(path)}" if len(parts) >= 3 else os.path.basename(path)


def _ghi(spec, path, recs):
    source_file = _source_id(path)
    conn = psycopg.connect(DB_URL)
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM raw_rows WHERE report_type=%s AND source_file=%s",
                    (spec["report_type"], source_file))
        by_ky, thieu_ky = {}, 0
        for r in recs:
            if not r.get("ngay"):
                thieu_ky += 1
                continue
            by_ky.setdefault(r["ngay"][:7], []).append(r)
        rows, skipped, i = [], [], 0
        base_i = int(spec.get("row_index_base", 6500000))
        for period, items in sorted(by_ky.items()):
            cur.execute("SELECT id FROM datasets WHERE kind='month' AND period=%s "
                        "ORDER BY created_at DESC LIMIT 1", (period,))
            got = cur.fetchone()
            if not got:
                skipped.append(period)
                continue
            for r in items:
                i += 1
                rows.append((got[0], spec["report_type"], base_i + i, r.get("ngay"),
                             r.get("cong_ty"), r.get("khoi"), r.get("cost_center"), period,
                             r.get("amount"), r.get("amount2"), r.get("dim1"), r.get("dim2"),
                             r.get("dim3"),
                             json.dumps(r.get("payload") or {}, ensure_ascii=False), source_file))
        if rows:
            cur.executemany(
                "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
                "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, "
                "source_file) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", rows)
        conn.commit()
        return {"written": len(rows),
                **({"bo_qua_chua_co_dataset": skipped} if skipped else {}),
                **({"bo_qua_khong_co_ngay": thieu_ky} if thieu_ky else {})}
    finally:
        conn.close()


def run(spec, path, write=False):
    recs, warn = extract_file(spec, path)
    by_ky = {}
    for r in recs:
        by_ky.setdefault((r.get("ngay") or "?")[:7], []).append(r)
    out = {"file": os.path.basename(path), "dong": len(recs),
           "ky": {k: {"dong": len(v),
                      "amount": round(sum(x.get("amount") or 0 for x in v), 6)}
                  for k, v in sorted(by_ky.items())}}
    if warn:
        out["canh_bao"] = warn
    # KHÔNG có bản ghi vẫn phải ghi (tức là XOÁ phần cũ của chính file này) khi file ĐÃ ĐỌC ĐƯỢC
    # mà bộ lọc loại hết dòng — đó là "file nguồn hết số", vd claim B2C T8 nhận 13/08/2026 có cột
    # Trạng thái + Số tiền toàn #N/A. Không xoá thì DB giữ nguyên bản nạp trước đó và dashboard
    # hiện số cũ như thể vẫn đúng.
    # Ngược lại, file KHÔNG đọc được (sai sheet, thiếu cột bắt buộc) trả rỗng mà KHÔNG kèm dòng bị
    # lọc -> giữ nguyên dữ liệu cũ: spec hỏng thì im lặng còn hơn xoá trắng dữ liệu đang đúng.
    if write and (recs or any(_W_BO_LOC in w for w in warn)):
        out.update(_ghi(spec, path, recs))
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Trích xuất báo cáo theo spec JSON.")
    ap.add_argument("spec", help="id spec (memory/extract_specs/<id>.json) hoặc đường dẫn file")
    ap.add_argument("--file", help="chỉ chạy 1 file (mặc định: quét cả folder trong spec)")
    ap.add_argument("--write", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    sp = load_spec(a.spec)
    if a.file:
        files, bo = [a.file], []
    else:
        files, w = quet_nguon(sp)
        for x in w:
            print(f"[!] {x}", file=sys.stderr)
        files, bo = loc_file_moi_nhat(sp, files)
    ket_qua = [run(sp, f, write=a.write) for f in files]
    if bo:
        ket_qua.append({"_bo_qua_anh_chup_cu": bo})
    print(json.dumps(ket_qua, ensure_ascii=False, indent=2))

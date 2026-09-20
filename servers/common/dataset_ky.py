# -*- coding: utf-8 -*-
"""KHAI SINH KỲ (`datasets` kind='month') — một chỗ duy nhất cho mọi đường ghi raw_rows.

VÌ SAO CÓ FILE NÀY (03/09/2026)
--------------------------------
`datasets` kind='month' KHÔNG phải "báo cáo tháng" — nó là CÁI KỲ. Mọi dòng `raw_rows` phải
treo vào một `dataset_id`, và mọi dòng thuộc kỳ 2026-09 (grain ngày lẫn tháng) đều treo chung
vào dataset kỳ 2026-09. Nhưng trước hôm nay chỉ ĐƯỜNG BÁO CÁO THÁNG (`template_filler`) mới tạo
được kỳ; ba đường ghi còn lại chỉ `SELECT`, không thấy thì bỏ:

  · `derive_hqkd_ngay`  -> trả lỗi "chưa có dataset tháng ..." và bỏ CẢ FILE;
  · `spec_extract`      -> bỏ LẶNG LẼ (chỉ nằm trong khoá `bo_qua_...` của JSON trả về);
  · `derive_kehoach_doanhthu` -> bỏ + báo `skipped`.

Hậu quả đo được ngày 03/09/2026: 5 file báo cáo NGÀY kỳ 2026-09 (Xanh Vĩnh Phúc, An Taxi, An
Khách Sạn, HTX Xanh Tuyên Quang, Dự án) đã về đĩa từ 16:36, deriver đọc ra số đàng hoàng (XVP
01/09 doanh thu 1,27 tỷ · 02/09 1,41 tỷ) nhưng không dòng nào vào được DB, chỉ vì kỳ 2026-09
chưa ai khai sinh. Tháng trước kỳ 2026-08 sinh ngày 04/08 -> đầu tháng nào cũng mất vài ngày,
năm nào cũng vậy, và không ai thấy vì báo cáo ngày lúc đó vẫn "đang chờ số".

NGUYÊN TẮC (chốt với người dùng 03/09/2026)
-------------------------------------------
Kỳ là hệ quả của SỐ THỰC TẾ ĐÃ PHÁT SINH, không phải của báo cáo tháng:

  · ai ghi số THỰC TẾ của một kỳ ĐÃ TỚI  -> được khai sinh kỳ đó (`tao=True`);
  · ai ghi KẾ HOẠCH / DỰ BÁO             -> KHÔNG (`tao=False`), chỉ điền vào kỳ đã có.

Dòng thứ hai không phải chi tiết vụn: `derive_kehoach_doanhthu` ghi TRỌN 12 KỲ trong một lần
nạp. Cho nó tạo kỳ thì ngay hôm nay dashboard mọc thêm 2026-10, 11, 12 — những kỳ chưa tồn tại,
chỉ có số kế hoạch, nằm chình ình trong ô chọn kỳ. Docstring của chính file đó đã chốt "không tự
tạo dataset rỗng" từ trước; nay lý do ấy nằm ở ĐÂY, một chỗ, thay vì là ba hành vi rời rạc tình
cờ khác nhau.

BỐN CHỐT AN TOÀN — đừng gỡ cái nào
-----------------------------------
1. CỬA SỔ KỲ HỢP LỆ: đúng dạng YYYY-MM và trong [tháng hiện tại − `LUI_TOI_DA` , tháng hiện
   tại] theo giờ VN. Chặn kỳ TƯƠNG LAI (bẫy kế hoạch ở trên) và kỳ RÁC do tên file gõ sai năm —
   có tiền lệ thật: file dầu Quang Sơn tiêu đề ghi 2025 trong khi là số 2026. Không có chốt này
   thì một cái typo đẻ ra một kỳ ma mà không ai lần được nguồn gốc.
2. KHÔNG BAO GIỜ `set_active`. Kỳ mới sinh ra với `is_active=0`. Kỳ đang xem chỉ đổi khi người
   dùng chọn, hoặc khi báo cáo tháng về qua `template_filler` (đường đó vẫn set_active như cũ,
   cố ý). Nếu không giữ chốt này thì đúng ngày 01 hằng tháng, một file báo cáo ngày có 1 ngày dữ
   liệu sẽ kéo cả dashboard nhảy sang tháng mới gần như trống.
3. KHOÁ `pg_advisory_xact_lock` THEO KỲ rồi SELECT LẠI trong cùng transaction. Bảng `datasets`
   KHÔNG có ràng buộc unique nào trên (kind, period) — hai lượt cron chạm cùng một kỳ trong cùng
   phút sẽ tạo hai dataset 2026-09; `resolve_id` lấy một bản, dòng gắn vào bản kia TÀNG HÌNH
   hoàn toàn, không báo gì. Prod và test lệch nhau 10 phút nhưng các job nguồn khác nhau vẫn có
   thể trùng phút.
4. NÓI TO CẢ HAI CHIỀU: hàm trả về trạng thái để chỗ gọi đưa `ky_moi_tao` và
   `bo_qua_ky_tuong_lai` lên JSON -> cron in ra log. Im lặng chính là thứ đã để chuyện này kéo
   dài hết tháng này sang tháng khác.

Dùng: nhận sẵn CURSOR của chỗ gọi (cả ba script đều tự mở psycopg) và ghi trong CHÍNH transaction
đó — chỗ gọi rollback thì kỳ vừa tạo cũng biến mất theo, đúng ý.
"""
import datetime
import re
import uuid

_RE_KY = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")

# Lùi tối đa 24 tháng: đủ cho mọi đợt nạp bù đã từng làm (xa nhất là 2025-09), mà vẫn chặn được
# typo lệch cả năm. Nạp bù xa hơn là việc CÓ NGƯỜI NGỒI CẠNH -> đi đường `template_filler`, nơi
# kỳ được tạo tường minh.
LUI_TOI_DA = 24

VN = datetime.timezone(datetime.timedelta(hours=7))

# Trạng thái trả về (phần tử thứ 2). "co_san"/"moi_tao" = có dataset_id; còn lại = None.
CO_SAN, MOI_TAO = "co_san", "moi_tao"
KY_KHONG_HOP_LE = "bo_qua_ky_khong_hop_le"
KY_TUONG_LAI = "bo_qua_ky_tuong_lai"
KY_QUA_CU = "bo_qua_ky_qua_cu"
KHONG_DUOC_TAO = "bo_qua_khong_duoc_tao"


def _thang_hien_tai() -> str:
    return f"{datetime.datetime.now(VN):%Y-%m}"


def _lui(period: str, n: int) -> str:
    y, m = int(period[:4]), int(period[5:7])
    i = y * 12 + (m - 1) - n
    return f"{i // 12:04d}-{i % 12 + 1:02d}"


def trong_cua_so(period: str) -> str:
    """'' nếu kỳ được phép khai sinh, ngược lại trả mã lý do."""
    if not _RE_KY.match(period or ""):
        return KY_KHONG_HOP_LE
    nay = _thang_hien_tai()
    if period > nay:
        return KY_TUONG_LAI
    if period < _lui(nay, LUI_TOI_DA):
        return KY_QUA_CU
    return ""


def lay_hoac_tao_ky(cur, period: str, *, nguon: str = "", tao: bool = True):
    """-> (dataset_id | None, trạng thái). `nguon` chỉ để ghi vết, không đổi hành vi.

    `tao=False` cho đường KẾ HOẠCH: tra cứu thôi, không khai sinh kỳ (xem docstring module).
    """
    if not _RE_KY.match(period or ""):
        return None, KY_KHONG_HOP_LE
    cur.execute("SELECT id FROM datasets WHERE kind='month' AND period=%s "
                "ORDER BY created_at DESC LIMIT 1", (period,))
    row = cur.fetchone()
    if row:
        return row[0], CO_SAN
    if not tao:
        return None, KHONG_DUOC_TAO
    ly_do = trong_cua_so(period)
    if ly_do:
        return None, ly_do
    # Chốt 3: khoá theo kỳ rồi HỎI LẠI — lượt song song có thể vừa tạo xong trong lúc ta chờ.
    cur.execute("SELECT pg_advisory_xact_lock(hashtext(%s))", (f"datasets:month:{period}",))
    cur.execute("SELECT id FROM datasets WHERE kind='month' AND period=%s "
                "ORDER BY created_at DESC LIMIT 1", (period,))
    row = cur.fetchone()
    if row:
        return row[0], CO_SAN
    ds_id = str(uuid.uuid4())
    # `created_at` là cột TEXT: giữ đúng khuôn UTC ISO của backend (dataset_repository.now_iso).
    # `is_active=0` — chốt 2.
    cur.execute(
        "INSERT INTO datasets (id, name, created_at, is_active, kind, period) "
        "VALUES (%s, %s, %s, 0, 'month', %s)",
        (ds_id, period,
         datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"), period))
    return ds_id, MOI_TAO


def giai_thich(period: str, trang_thai: str) -> str:
    """Câu giải thích cho log/UI khi không lấy được kỳ."""
    return {
        KY_KHONG_HOP_LE: f"kỳ '{period}' không đúng dạng YYYY-MM — kiểm tra tên file nguồn",
        KY_TUONG_LAI: (f"kỳ {period} CHƯA TỚI nên không tự tạo (chỉ số thực tế mới được khai "
                       f"sinh kỳ; kế hoạch/dự báo thì không) — số của kỳ này chưa được nạp"),
        KY_QUA_CU: (f"kỳ {period} cũ hơn {LUI_TOI_DA} tháng nên không tự tạo — nghi tên file ghi "
                    f"sai năm; muốn nạp bù thật thì tạo kỳ tường minh rồi chạy lại"),
        KHONG_DUOC_TAO: (f"chưa có kỳ {period} và nguồn này (kế hoạch/dự báo) không được phép "
                         f"khai sinh kỳ — chờ số thực tế của kỳ đó về trước"),
    }.get(trang_thai, f"không lấy được kỳ {period} ({trang_thai})")

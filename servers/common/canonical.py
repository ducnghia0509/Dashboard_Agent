# -*- coding: utf-8 -*-
"""Bang alias RUT GON cho 'loai bao cao chuan hoa' (canonical_kind) - dung de nhan
dien 1 sheet/file thuoc loai bao cao ke toan nao BAT KE ten sheet/file that su dat
la gi (vd '131' hay 'Cong no phai thu' deu la TK131).

Co tinh: KHONG tao alias.json rieng (tranh nhan ban nguon su that voi kpi_glossary.json/
FIELD_DEFS da co) - day chi la 1 dict nho, deterministic, sua truc tiep trong code khi
gap loai bao cao moi."""
from . import be_bridge as bb

# Thứ tự khai báo = thứ tự ưu tiên khớp (guess trả kind đầu tiên trúng alias). Alias phải
# ĐỦ ĐẶC TRƯNG để không đè nhau — đã kiểm với tên sheet thật của BaocaoHQKD (1-3.HQKD, 4.doanh
# thu, 5.Bán xe, 6.Pthu, 7.Ptra, 8.Đầu tư, 9.hàng hóa, 10.Tài sản) và TC01 (SD TIỀN, BC THU CHI).
# Khớp theo TÊN sheet (tên ngắn, do người đặt có chủ đích: '131','Pthu','9. hàng hóa'...).
CANONICAL_KINDS = {
    "KQKD": ["kqkd", "hqkd", "bckqkd", "kqhdkd", "bc kqhdkd", "bao cao ket qua kinh doanh", "ket qua hoat dong kinh doanh"],
    "CDKT": ["cdkt", "bang can doi ke toan", "can doi ke toan"],
    # B32: "cdsps" = viết tắt "Cân Đối Số Phát Sinh" (CĐSPS, 5 chữ — CÓ "Số") — khác "cdps"
    # (Cân Đối Phát Sinh, 4 chữ — THIẾU "Số"). 2 cách viết tắt cùng 1 khái niệm, một số công ty
    # dùng cách này, một số dùng cách khác (xem _chung.yaml "sheet_nguon: CDPS / CĐSPS") — thiếu
    # "cdsps" khiến sheet đặt tên "CĐSPS" không nhận diện được gì cả (test xác nhận: None).
    "CDPS": ["cdps", "cdsps", "can doi so phat sinh", "bang can doi so phat sinh"],
    "TK131": ["131", "tk131", "pthu", "phai thu", "cong no phai thu"],
    "TK331": ["331", "tk331", "ptra", "phai tra", "cong no phai tra"],
    "TSCD": ["tscd", "tai san co dinh", "tai san", "khau hao", "bieu khau hao", "khts", "khcd"],
    "LCTT": ["lctt", "luu chuyen tien te", "bao cao luu chuyen tien te"],
    "SODU_TIEN": ["sd tien", "so du tien", "sodu tien"],   # kiểm TRƯỚC THUCHI (đều có 'tien')
    "THUCHI": ["bc thu chi", "thu chi", "dong tien"],
    "VAY": ["khoan vay", "no vay", "du no vay"],           # KHÔNG để trơ "vay" (dễ khớp nhầm)
    "TONKHO": ["ton kho", "hang ton kho", "hang hoa", "nhap xuat ton", "nxt", "156"],
    "THUE": ["thue"],
    "DAUTU": ["dau tu"],
    "KDVH": ["ban xe", "kdvh"],
}

# Khớp theo NỘI DUNG (tiêu đề trong sheet) khi TÊN sheet khó đọc ('156','KHTS','Sheet1').
# CHỈ dùng cụm ĐẶC TRƯNG >=3 từ — KHÔNG dùng token ngắn ('thue','phai tra','tai san','131'):
# test cho thấy token ngắn khớp lung tung vào ô/header -> route NHẦM (PTB2B/DATA/211+242 -> sai).
# Quét trên CỘT TIÊU ĐỀ (ô đầu mỗi dòng), không phải toàn bộ lưới.
_CONTENT_KINDS = {
    "KQKD": ["ket qua kinh doanh", "ket qua hoat dong kinh doanh", "bao cao loi nhuan", "bao cao lai lo"],
    "CDKT": ["bang can doi ke toan"],
    "CDPS": ["bang can doi so phat sinh", "can doi so phat sinh"],
    "TK131": ["cong no phai thu"],
    "TK331": ["cong no phai tra"],
    "TSCD": ["bang tinh khau hao", "khau hao tai san co dinh", "khau hao tscd"],
    "TONKHO": ["nhap xuat ton", "nhap - xuat - ton", "bao cao ton kho"],
    "LCTT": ["luu chuyen tien te"],
    "SODU_TIEN": ["bao cao so du tien"],
    "VAY": ["tinh hinh vay", "du no vay"],
}


def guess_kind_from_content(title_text: str):
    """Khớp canonical_kind theo TIÊU ĐỀ báo cáo (cụm đặc trưng, strict). None nếu không chắc
    -> caller để 'unknown' (đẩy analyst/người) thay vì route nhầm."""
    if not title_text:
        return None
    norm = bb.remove_diacritics(bb.normalize_header(title_text))
    for kind, phrases in _CONTENT_KINDS.items():
        if any(p in norm for p in phrases):
            return kind
    return None

# Sheet metadata/tổng hợp KHÔNG phải dữ liệu nhập -> bỏ khi auto-route (tránh nạp nhầm/đúp).
# '(tong)': sheet 'BC THU CHI (TỔNG)' là tổng của các sheet công ty -> nạp sẽ đúp số.
_SKIP_SHEET_ALIASES = ["master data", "ds bao cao", "danh sach bao cao", "user",
                       "(tong)", "huong dan", "muc luc"]


def guess_canonical_kind(text: str):
    """text: tên sheet hoặc tên file. Trả về 1 canonical_kind khớp đầu tiên, hoặc None.
    Match theo substring sau khi bỏ dấu + hạ chữ thường (KHÔNG fuzzy) - đủ dùng vì
    alias list đã liệt kê rõ các biến thể hay gặp; fuzzy hơn thì dùng glossary_lookup."""
    norm = bb.remove_diacritics(bb.normalize_header(text))
    # GUARD sheet THUẾ: tên có 'thue' + dạng công nợ ('phai tra/thu/nop') -> THUE, KHÔNG để TK331/
    # TK131 (alias 'phai tra'/'phai thu') khớp trước. Nếu không guard, sheet 'Thuế phải trả - phải
    # nộp' route TK331 -> cùng target 06_PHAITRA với sheet 'Phải trả 331' thật; import_filled
    # delete-scope theo source_file khiến sheet thuế (chạy sau) ĐÈ MẤT toàn bộ công nợ NCC thật
    # (XVP T06: 59 dòng NCC -> còn 7 dòng thuế = 0). Chỉ khớp khi CÓ CẢ 'thue' lẫn token công nợ
    # -> không đụng CĐPS/CĐKT/KQKD (tên không chứa đồng thời 2 cụm này).
    if "thue" in norm and any(k in norm for k in ("phai tra", "phai thu", "phai nop")):
        return "THUE"
    for kind, aliases in CANONICAL_KINDS.items():
        if any(a in norm for a in aliases):
            return kind
    return None


def is_skip_sheet(name: str) -> bool:
    """True nếu sheet là metadata/tổng hợp (Master Data, User, DS báo cáo, BC THU CHI (TỔNG)...)
    — auto-route BỎ QUA, không coi là sheet dữ liệu để nạp."""
    norm = bb.remove_diacritics(bb.normalize_header(name))
    return any(a in norm for a in _SKIP_SHEET_ALIASES)


import re as _re


def _slug_gen(text: str) -> str:
    """Chuẩn hoá 1 chuỗi tự do về slug ổn định (HOA, bỏ dấu, gom mọi ký tự lạ về '_')."""
    slug = _re.sub(r"[^A-Z0-9]+", "_", bb.remove_diacritics(str(text or "")).upper()).strip("_")
    return slug


def canonical_gen_code(raw: str) -> str:
    """Chuẩn hoá 1 report_type GENERIC về dạng GEN_<canonical> ỔN ĐỊNH giữa các tháng/công ty.

    Vì report_type generic (target_report_type) do LLM/heuristic ĐẶT TỰ DO nên cùng 1 khái niệm
    có thể ra nhiều chuỗi khác nhau (vd 'GEN_10THUE' vs 'GEN_10_THUE' cùng là sheet CĐPS/thuế;
    'GEN_TK331' vs 'GEN_PTRA' cùng là phải trả). Điều đó làm tập chiều lệch theo tháng và phá
    idempotent (2 tên khác nhau -> không xoá đè nhau -> nạp đúp). Hàm này gom mọi biến thể về 1
    mã duy nhất bằng CHÍNH bảng alias canonical (single source of truth): bỏ tiền tố 'GEN_', map
    phần còn lại về canonical_kind qua guess_canonical_kind -> 'GEN_<KIND>'. Không map được thì
    giữ nguyên nội dung nhưng chuẩn hoá chuỗi (slug) để 2 cách viết cùng khái niệm hội tụ.

    Idempotent: canonical_gen_code('GEN_THUE') == 'GEN_THUE'.
    """
    s = (raw or "").strip()
    if not s:
        return "GEN_UNKNOWN"
    if s.upper().startswith("GEN_"):
        s = s[4:]
    kind = guess_canonical_kind(s)
    if kind:
        return f"GEN_{kind}"
    slug = _slug_gen(s)
    return f"GEN_{slug}" if slug else "GEN_UNKNOWN"


# ---- Extension 3: scope (hợp nhất/riêng) + basis (lũy kế/theo kỳ) — thuyết minh ngắn ----
# Cùng triết lý CANONICAL_KINDS: dict nhỏ, deterministic, KHÔNG fuzzy. Đọc từ tên file +
# text trong sheet (row_sample đã có sẵn từ sheet_profile, không mở thêm file).

_SCOPE_ALIASES = {
    "hopnhat": ["hop nhat", "consol"],
    "rieng": ["rieng"],  # "riêng"/"baocaotaichinhrieng" đều chứa substring này sau bỏ dấu
}
_BASIS_ALIASES = {
    "luyke": ["luy ke", "tu dau nam", "luy ke tu dau nam"],
    "thang": ["tu ngay", "trong thang", "trong ky"],  # có khoảng ngày cụ thể trong 1 kỳ -> theo kỳ
}


def _match_alias(text: str, table: dict):
    if not text:
        return None
    norm = bb.remove_diacritics(bb.normalize_header(text))
    for key, aliases in table.items():
        if any(a in norm for a in aliases):
            return key
    return None


def guess_scope(text: str):
    """'hopnhat' | 'rieng' | None. text: tên file hoặc text tiêu đề sheet (row_sample)."""
    return _match_alias(text, _SCOPE_ALIASES)


def guess_basis(text: str):
    """'luyke' | 'thang' | None. text: text tiêu đề sheet (vd 'Từ ngày...Đến ngày...' -> thang,
    'Lũy kế từ đầu năm' -> luyke). KHÔNG suy từ tên file (basis hiếm khi có trong tên file)."""
    return _match_alias(text, _BASIS_ALIASES)


# nhom_bao_cao/nhom_con (kpi_glossary.json, từ guideline.xlsx) -> screen id trong src/nav.ts.
# Alias nhỏ, best-effort — dùng để ĐỀ XUẤT màn hình liên quan, KHÔNG tự động ghi số vào đó.
_SCREEN_BY_NHOM_CON = {
    "congno": ["phai thu", "phai tra", "cong no"],
    "tonkho": ["ton kho", "hang hoa"],
    "taisan": ["tai san co dinh", "khau hao"],
    "dautu": ["dau tu"],
}
_SCREEN_BY_NHOM_BAO_CAO = {
    "dongtien": ["dong tien"],
    "hieuqua": ["hieu qua kinh doanh"],
    "taisan-nguonvon": ["tai san - nguon von", "tai san nguon von"],
}


def guess_screen(nhom_bao_cao: str = "", nhom_con: str = ""):
    """Đề xuất screen id (khớp src/nav.ts) từ nhom_bao_cao/nhom_con trong kpi_glossary.json.
    Ưu tiên nhom_con (cụ thể hơn) trước, fallback nhom_bao_cao (nhóm lớn). None nếu không khớp
    — không đoán liều khi guideline không đủ rõ."""
    screen = _match_alias(nhom_con, _SCREEN_BY_NHOM_CON)
    if screen:
        return screen
    return _match_alias(nhom_bao_cao, _SCREEN_BY_NHOM_BAO_CAO)

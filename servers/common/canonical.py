# -*- coding: utf-8 -*-
"""Bang alias RUT GON cho 'loai bao cao chuan hoa' (canonical_kind) - dung de nhan
dien 1 sheet/file thuoc loai bao cao ke toan nao BAT KE ten sheet/file that su dat
la gi (vd '131' hay 'Cong no phai thu' deu la TK131).

Co tinh: KHONG tao alias.json rieng (tranh nhan ban nguon su that voi kpi_glossary.json/
FIELD_DEFS da co) - day chi la 1 dict nho, deterministic, sua truc tiep trong code khi
gap loai bao cao moi."""
from . import be_bridge as bb

CANONICAL_KINDS = {
    "KQKD": ["kqkd", "bckqkd", "bao cao ket qua kinh doanh", "ket qua hoat dong kinh doanh"],
    "CDKT": ["cdkt", "bang can doi ke toan", "can doi ke toan"],
    "CDPS": ["cdps", "can doi so phat sinh", "bang can doi so phat sinh"],
    "TK131": ["131", "tk131", "phai thu", "cong no phai thu"],
    "TK331": ["331", "tk331", "phai tra", "cong no phai tra"],
    "TSCD": ["tscd", "tai san co dinh", "khau hao", "bieu khau hao"],
    "LCTT": ["lctt", "luu chuyen tien te", "bao cao luu chuyen tien te"],
}


def guess_canonical_kind(text: str):
    """text: tên sheet hoặc tên file. Trả về 1 canonical_kind khớp đầu tiên, hoặc None.
    Match theo substring sau khi bỏ dấu + hạ chữ thường (KHÔNG fuzzy) - đủ dùng vì
    alias list đã liệt kê rõ các biến thể hay gặp; fuzzy hơn thì dùng glossary_lookup."""
    norm = bb.remove_diacritics(bb.normalize_header(text))
    for kind, aliases in CANONICAL_KINDS.items():
        if any(a in norm for a in aliases):
            return kind
    return None


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

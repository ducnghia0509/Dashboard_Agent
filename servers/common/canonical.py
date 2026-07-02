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

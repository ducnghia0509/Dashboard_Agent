# -*- coding: utf-8 -*-
"""DANH MỤC TỔ CHỨC — nguồn define DUY NHẤT cho filler (cấp độ / công ty / khối / cost center).

Dữ liệu ở org_catalog.json (sinh bởi scripts/gen_org.py, seed từ Template_chuan MD_ đang chạy).
SỬA org_catalog.json để mở rộng — loader này đọc lại theo mtime (không cần restart).
Thay cho các map hard-code rải rác trước đây (_CO trong extractor, _PATH_KHOI trong contract).
"""
import json
import os
import unicodedata

_HERE = os.path.dirname(os.path.abspath(__file__))
_CATALOG = os.path.join(_HERE, "org_catalog.json")
_cache = None
_cache_mtime = None


def _norm(s: str) -> str:
    s = str(s or "").lower().replace("đ", "d").strip()
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def load() -> dict:
    """Nạp (cache theo mtime) org_catalog.json. {} nếu thiếu file (an toàn)."""
    global _cache, _cache_mtime
    try:
        mtime = os.path.getmtime(_CATALOG)
    except OSError:
        return _cache or {}
    if _cache is not None and _cache_mtime == mtime:
        return _cache
    with open(_CATALOG, encoding="utf-8") as fh:
        _cache = json.load(fh)
    _cache_mtime = mtime
    return _cache


# ── CÔNG TY ──────────────────────────────────────────────────────────────────────
def companies() -> list:
    return load().get("companies", [])


_FUZZY_MIN_LEN = 4   # B31: token < 4 ký tự (vd "HO") gần như chắc khớp NHẦM vào chuỗi con bất kỳ
                     # (đã bắt thực tế: "HO" in "Hợp tác xã..."/"Showroom" -> sai hoàn toàn) —
                     # dưới ngưỡng này CHỈ nhận khớp mã CHÍNH XÁC, không containment-match.


def company_code(name_or_code) -> str | None:
    """Tên (đầy đủ/rút gọn/alias) hoặc mã -> MÃ công ty chuẩn. None nếu không khớp."""
    if not name_or_code:
        return None
    n = _norm(name_or_code)
    comps = companies()
    for c in comps:                                   # 1) khớp mã trực tiếp (luôn cho phép)
        if _norm(c["ma"]) == n:
            return c["ma"]
    if len(n) < _FUZZY_MIN_LEN:
        return None
    for c in comps:                                   # 2) khớp alias (chứa)
        if any(a and (_norm(a) in n or n in _norm(a)) for a in c.get("aliases", [])):
            return c["ma"]
    for c in comps:                                   # 3) khớp tên đầy đủ (chứa)
        if _norm(c["ten"]) and (n in _norm(c["ten"]) or _norm(c["ten"]) in n):
            return c["ma"]
    return None


def company_name(code) -> str | None:
    return next((c["ten"] for c in companies() if c["ma"] == code), None)


# ── KHỐI ─────────────────────────────────────────────────────────────────────────
def khoi() -> list:
    return load().get("khoi", [])


def khoi_name(code) -> str | None:
    return next((k["ten"] for k in khoi() if str(k["ma"]) == str(code)), None)


def khoi_code(name) -> str | None:
    n = _norm(name)
    for k in khoi():                                  # 1) khớp CHÍNH XÁC (luôn cho phép)
        if _norm(k["ten"]) == n:
            return k["ma"]
    if len(n) < _FUZZY_MIN_LEN:                        # B31: token ngắn (vd "HO") -> không containment
        return None
    for k in khoi():                                  # 2) khớp chứa (chỉ khi token đủ dài)
        if n in _norm(k["ten"]) or _norm(k["ten"]) in n:
            return k["ma"]
    return None


def khoi_names() -> list:
    return [k["ten"] for k in khoi()]


def khoi_from_path(path: str) -> str | None:
    """Suy TÊN khối từ đường dẫn nguồn: khớp CHÍNH XÁC từng segment thư mục với path_khoi.
    (vd .../SRVF/... -> 'Khối KD Xe điện Vinfast - SR'). None nếu không khớp.

    CHỈ nên dùng làm FALLBACK cho công ty ĐA KHỐI (xem khoi_for_company) — thư mục lưu file có
    thể là thư mục HÀNH CHÍNH/tổ chức (vd file công ty GA lại nộp vào thư mục 'HO') KHÔNG phản
    ánh đúng khối công ty đó vận hành, nên KHÔNG được ưu tiên hơn mã công ty khi công ty đó chỉ
    thuộc đúng 1 khối."""
    if not path:
        return None
    import re
    pk = load().get("path_khoi", {})
    segs = [s.strip().upper().replace(" ", "").replace("_", "")
            for s in re.split(r"[\\/]+", str(path)) if s.strip()]
    for s in segs:
        if s in pk:
            return khoi_name(pk[s]) or pk[s]
    return None


_KNOWLEDGE_DIR = os.environ.get("KNOWLEDGE_DIR", "/home/sysadmin/knowledge")
_MAP_PATH = os.path.join(_KNOWLEDGE_DIR, "khoi_phapnhan_map.yaml")
_map_cache = None
_map_cache_mtime = None


def _load_khoi_phapnhan_map() -> dict:
    """{mã công ty -> tập ma_khoi} suy từ khoi_phapnhan_map.yaml (quan_he[].phap_nhan_hqkd) —
    quy tắc NGHIỆP VỤ THẬT (không suy từ thư mục lưu file). Nạp lại theo mtime; {} nếu thiếu file."""
    global _map_cache, _map_cache_mtime
    try:
        mtime = os.path.getmtime(_MAP_PATH)
    except OSError:
        return _map_cache or {}
    if _map_cache is not None and _map_cache_mtime == mtime:
        return _map_cache
    import yaml
    with open(_MAP_PATH, encoding="utf-8") as fh:
        doc = yaml.safe_load(fh) or {}
    idx: dict = {}
    for rel in doc.get("quan_he", []) or []:
        ma_khoi = str(rel.get("ma_khoi") or "").strip()
        if not ma_khoi:
            continue
        for cty in (rel.get("phap_nhan_hqkd") or []):
            idx.setdefault(str(cty).strip().upper(), set()).add(ma_khoi)
    _map_cache, _map_cache_mtime = idx, mtime
    return idx


def khoi_for_company(company_code: str) -> str | None:
    """TÊN khối DUY NHẤT của công ty theo khoi_phapnhan_map.yaml (nguồn nghiệp vụ THẬT, ưu tiên
    hơn suy từ đường dẫn lưu file). Trả None nếu công ty không xác định/không có trong map, HOẶC
    công ty vận hành ĐA KHỐI (TC/XVP/AAG — khi đó phải suy thêm theo path/cost center của TỪNG
    file, xem khoi_from_path) — tức 1 mã công ty không đủ để chốt 1 khối."""
    if not company_code:
        return None
    khois = _load_khoi_phapnhan_map().get(str(company_code).strip().upper())
    if not khois or len(khois) != 1:
        return None
    ma = next(iter(khois))
    return khoi_name(ma) or ma


# ── COST CENTER ──────────────────────────────────────────────────────────────────
def cost_centers() -> list:
    return load().get("cost_centers", [])


def cost_center(ma) -> dict | None:
    return next((cc for cc in cost_centers() if cc["ma"] == ma), None)


# ── CHIỀU BỔ SUNG ────────────────────────────────────────────────────────────────
def data_levels() -> dict:
    return load().get("data_levels", {})


def frequencies() -> dict:
    return load().get("frequencies", {})

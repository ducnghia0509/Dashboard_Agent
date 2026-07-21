# -*- coding: utf-8 -*-
"""Lớp 2 — HƯỚNG DẪN LẤY DỮ LIỆU per-đơn vị (xem knowledge/00_index.yaml).

Đọc /home/sysadmin/knowledge/{ma_congty}.yaml NẾU CÓ (thay thế hoàn toàn, KHÔNG merge), ngược lại
đọc knowledge/_chung.yaml. Đây là loader MỎNG — nội dung/luật nghiệp vụ do business viết trong YAML,
thêm 1 công ty có quirk riêng chỉ cần tạo file mới, không đụng code này.

Cache theo mtime (giống org.py/contract.py) — sửa yaml có hiệu lực ngay, không cần restart service.
"""
import glob
import os

import yaml

KNOWLEDGE_DIR = os.environ.get("KNOWLEDGE_DIR", "/home/sysadmin/knowledge")
DEFAULT_GUIDE = "_chung.yaml"

# Tên file HỆ THỐNG trong knowledge/ — KHÔNG BAO GIỜ coi là hướng dẫn per-công-ty, dù trùng basename
# với 1 ma_congty tương lai (an toàn cho trường hợp trùng tên khó xảy ra).
_RESERVED = {
    "companies.yaml", "khoi.yaml", "cost_centers.yaml", "khoi_phapnhan_map.yaml",
    "source_path_rules.yaml", "template_columns.yaml", "screen_sources.yaml",
    "classification.yaml", "screen_endpoints.yaml", "00_index.yaml", DEFAULT_GUIDE,
}

_cache = {}  # path -> (mtime, content)


def _read(path):
    mtime = os.path.getmtime(path)
    cached = _cache.get(path)
    if cached and cached[0] == mtime:
        return cached[1]
    with open(path, encoding="utf-8") as fh:
        content = yaml.safe_load(fh)
    _cache[path] = (mtime, content)
    return content


def load_guide(ma_congty: str = None, file_name: str = None) -> dict:
    """Trả {"source", "ma_congty", "is_specific", "content"} (+ "error" nếu thiếu cả 2 file).

    Luật load (00_index.yaml):
      - ma_congty đã biết + tồn tại knowledge/{ma_congty}.yaml -> đọc file đó, THAY THẾ hoàn toàn.
      - Ngược lại (không có file riêng, hoặc chưa xác định được ma_congty) -> đọc _chung.yaml.

    file_name: nếu ma_congty chưa biết, thử suy từ tên file (delegate contract.resolve_company)
    để có cơ hội tìm đúng hướng dẫn riêng ngay từ lần gọi đầu, không cần caller tự suy trước.
    """
    # CHUẨN HOÁ ma_congty về MÃ CÔNG TY HỢP LỆ (trong companies.yaml/md_congty). ma_congty caller
    # truyền có thể là TÊN THƯ MỤC/KHỐI lấy nhầm từ path/sidecar ('HO','DUAN','ANTAXI'...) — không
    # phải mã cty. resolve_company BỎ raw không hợp lệ rồi suy mã đúng từ tên file ('B.8.GA.' -> GA);
    # raw HỢP LỆ vẫn được ưu tiên giữ. Nhờ vậy dù caller đưa 'HO' vẫn tìm được GA.yaml theo tên file.
    if ma_congty or file_name:
        try:
            from . import contract
            resolved = contract.resolve_company(ma_congty, file_name)
            if resolved:
                ma_congty = resolved
            elif ma_congty and str(ma_congty).strip().upper() not in {
                    str(k).strip().upper() for k in contract.load_contract()["md_congty"]}:
                ma_congty = None  # 'HO'/khối không hợp lệ & tên file cũng không cho mã -> _chung.yaml
        except Exception:
            pass

    ma = str(ma_congty).strip() if ma_congty else None
    if ma:
        # BIẾN THỂ theo khối/layout riêng (vd TC_SRVF.yaml cho ma_congty=TC khi file có "SRVF") —
        # khớp nếu tên file chứa token sau "{ma}_" của 1 file {ma}_<TOKEN>.yaml đang có trong
        # knowledge/. Không hard-code tên biến thể cụ thể trong code — thêm file mới là đủ.
        if file_name:
            variants = [p for p in sorted(glob.glob(os.path.join(KNOWLEDGE_DIR, f"{ma}_*.yaml")))
                        if os.path.basename(p) not in _RESERVED]
            # (1) KHỚP THEO MÃ KHỐI — tin cậy nhất: tên file 'B.<khối>.<mã cty>.' (vd 'B.3.TC.') so
            # với don_vi.ma_khoi của từng guide. Thay khớp token-chuỗi cũ vì nó vừa BỎ SÓT (basename
            # 'B.3.TC...' KHÔNG chứa 'TRAMSAC' -> guide riêng không bao giờ được chọn) vừa KHỚP NHẦM
            # (token ngắn 'HO' trùng '/home/' trong đường dẫn -> chọn nhầm TC_HO).
            import re as _re
            mkh = _re.search(r"\bB\.?(\d{1,2})\.", str(file_name))
            file_khoi = mkh.group(1) if mkh else None
            if file_khoi:
                for vp in variants:
                    try:
                        dv = (_read(vp) or {}).get("don_vi", {}) or {}
                        if str(dv.get("ma_khoi", "")).strip() == file_khoi:
                            return {"source": os.path.basename(vp), "ma_congty": ma,
                                    "is_specific": True, "content": _read(vp)}
                    except Exception:
                        pass
            # (2) FALLBACK: token trong tên file — CHỈ token đủ dài (>=3) để tránh khớp nhầm token ngắn.
            for vp in variants:
                token = os.path.basename(vp)[len(ma) + 1:-len(".yaml")]
                if len(token) >= 3 and token.lower() in str(file_name).lower():
                    return {"source": os.path.basename(vp), "ma_congty": ma, "is_specific": True,
                            "content": _read(vp)}

        fname = f"{ma}.yaml"
        path = os.path.join(KNOWLEDGE_DIR, fname)
        if fname not in _RESERVED and os.path.exists(path):
            return {"source": fname, "ma_congty": ma, "is_specific": True, "content": _read(path)}

    default_path = os.path.join(KNOWLEDGE_DIR, DEFAULT_GUIDE)
    if not os.path.exists(default_path):
        return {"source": None, "ma_congty": ma, "is_specific": False, "content": None,
                "error": f"Không tìm thấy {DEFAULT_GUIDE} trong {KNOWLEDGE_DIR}"}
    return {"source": DEFAULT_GUIDE, "ma_congty": ma, "is_specific": False,
            "content": _read(default_path)}

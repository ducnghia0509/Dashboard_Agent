# -*- coding: utf-8 -*-
"""SINH org_catalog.json — DANH MỤC TỔ CHỨC (artifact sinh ra, KHÔNG sửa tay).

NGUỒN SỰ THẬT = /home/sysadmin/knowledge/*.yaml (companies/khoi/cost_centers). Sửa YAML rồi chạy
lại (hoặc knowledge/build.py). aliases/path_khoi ưu tiên đọc từ YAML nếu có field, không thì dùng
fallback hardcode bên dưới (companies.yaml/source_path_rules.yaml hiện là bản thuần nghiệp vụ,
không mang field kỹ thuật). Cấp độ dữ liệu (B/F/H/T) + tần suất (D/W/M/Q/Y) lấy từ Danh_Muc_Ma_he_thong.xlsx.

Giữ NGUYÊN shape org_catalog.json cũ -> gen_fe_master.py / org.py / master.py KHÔNG phải đổi.
Công ty "aggregate-only" (vd GR — mã hợp nhất, không phải pháp nhân thật) bị LOẠI khỏi danh mục:
suy TỪ DỮ LIỆU (mã không có cost center nào gán trong cost_centers.yaml), không cần cờ role thủ công.

Chạy:
  .venv/bin/python scripts/gen_org.py                 # ghi servers/common/org_catalog.json
  .venv/bin/python scripts/gen_org.py --out /tmp/x.json  # ghi ra bản COPY để review/diff
"""
import json
import os
import sys

import openpyxl
import yaml

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
KNOWLEDGE = os.environ.get("KNOWLEDGE_DIR", "/home/sysadmin/knowledge")
DANHMUC = os.path.join(_ROOT, "context", "sources", "Danh_Muc_Ma_he_thong.xlsx")
OUT = os.path.join(_ROOT, "servers", "common", "org_catalog.json")

# Fallback nếu YAML thiếu (an toàn khi chuyển tiếp). Nguồn chính giờ là YAML.
_ALIASES_FALLBACK = {
    "TC": ["thịnh cường", "tc"],
    "VFQN": ["vinfast quảng ninh", "vfqn", "công nghệ vinfast quảng ninh"],
    "GA": ["global ai", "ga"],
    "AAG": ["an an", "an taxi", "an an's garden", "an ks", "aag"],
    "XVP": ["xanh vĩnh phúc", "xanh vp", "xvp"],
    "HTX_XVP": ["hợp tác xã vận tải xanh vĩnh phúc", "htx vĩnh phúc"],
    "HTX_XTQ": ["hợp tác xã vận tải xanh tuyên quang", "htx tuyên quang"],
    "HT": ["hưng thịnh", "ht"],
}
# Cập nhật 2026-07-08 theo trust_me_bro.xlsx — mã khối đánh số lại toàn bộ (xem knowledge/khoi.yaml).
# ANTAXI nay trỏ khối 7 "Dịch vụ An Taxi" (tách riêng khỏi khối Taxi Xanh, không còn dùng chung mã
# với XANHVINHPHUC). source_path_rules.yaml không có key 'path_khoi' (chỉ có 'patterns' cho người
# đọc) -> _path_khoi() luôn dùng fallback này; SỬA Ở ĐÂY là sửa nguồn thật.
_PATH_KHOI_FALLBACK = {
    "SRVF": "1", "XDV": "2", "TRAMSAC": "3", "DUAN": "4", "HO": "9",
    "XANHVINHPHUC": "6", "ANTAXI": "7", "ANKHACHSAN": "10", "HUNGTHINH": "5",
}


def _load_yaml(name):
    path = os.path.join(KNOWLEDGE, name)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Thiếu YAML nguồn: {path} (knowledge là nguồn sự thật).")
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def _companies():
    """companies.yaml -> [{ma,ten,aliases}], LOẠI mã "aggregate-only" (vd GR = mã hợp nhất,
    không phải pháp nhân vận hành thật). Suy TỪ DỮ LIỆU: mã KHÔNG có bất kỳ cost center nào
    gán trong cost_centers.yaml -> loại. Không cần business gắn cờ kỹ thuật (role) vào companies.yaml
    — companies.yaml chỉ cần mô tả nghiệp vụ thuần (ma_congty/ten_congty/thuyet_minh)."""
    used = {str(cc["ma_congty"]).strip() for cc in (_load_yaml("cost_centers.yaml") or [])
            if cc.get("ma_congty")}
    out = []
    for c in _load_yaml("companies.yaml") or []:
        ma = str(c["ma_congty"]).strip()
        if ma not in used:
            continue
        aliases = c.get("aliases") or _ALIASES_FALLBACK.get(ma, [ma.lower()])
        out.append({"ma": ma, "ten": str(c["ten_congty"]).strip(), "aliases": list(aliases)})
    return out


def _khoi():
    return [{"ma": str(k["ma_khoi"]).strip(), "ten": str(k["ten_khoi"]).strip()}
            for k in (_load_yaml("khoi.yaml") or [])]


def _cost_centers():
    out = []
    for cc in _load_yaml("cost_centers.yaml") or []:
        if not cc.get("ma_cost_center"):
            continue
        out.append({
            "ma": str(cc["ma_cost_center"]).strip(),
            "ten": str(cc.get("ten_cost_center") or "").strip(),
            "cong_ty": str(cc["ma_congty"]).strip() if cc.get("ma_congty") else None,
            "khoi": str(cc["ma_khoi"]).strip() if cc.get("ma_khoi") is not None else None,
        })
    return out


def _path_khoi():
    try:
        spr = _load_yaml("source_path_rules.yaml") or {}
    except FileNotFoundError:
        return dict(_PATH_KHOI_FALLBACK)
    pk = spr.get("path_khoi") if isinstance(spr, dict) else None
    return {str(k).strip().upper(): str(v).strip() for k, v in (pk.items() if pk else _PATH_KHOI_FALLBACK.items())}


def _extra_dims():
    """Cấp độ dữ liệu + tần suất báo cáo — lấy từ Danh_Muc (chưa đưa vào YAML)."""
    levels, freq = {}, {}
    if not os.path.exists(DANHMUC):
        return levels, freq
    wb = openpyxl.load_workbook(DANHMUC, data_only=True, read_only=True)
    ws = wb["Sheet1"]
    rows = [[("" if c is None else str(c).strip()) for c in r] for r in ws.iter_rows(values_only=True)]
    for r in rows[3:12]:
        if len(r) > 2 and r[2] and len(r[2]) == 1 and r[1]:
            levels[r[2]] = r[1]
        if len(r) > 17 and r[17] and r[16]:
            freq[r[17]] = r[16]
    wb.close()
    return levels, freq


def build(out_path=None):
    companies = _companies()
    khoi = _khoi()
    ccs = _cost_centers()
    path_khoi = _path_khoi()
    levels, freq = _extra_dims()
    catalog = {
        "_source": "SINH TỪ knowledge/*.yaml (companies/khoi/cost_centers + aliases + source_path_rules). "
                   "KHÔNG sửa tay file này. Cấp độ/tần suất từ Danh_Muc. Chạy knowledge/build.py để đồng bộ.",
        "data_levels": levels,
        "frequencies": freq,
        "khoi": khoi,
        "companies": companies,
        "cost_centers": ccs,
        "path_khoi": path_khoi,
    }
    dest = out_path or OUT
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(dest, "w", encoding="utf-8") as fh:
        json.dump(catalog, fh, ensure_ascii=False, indent=2)
    return {"ok": True, "out": dest, "khoi": len(khoi), "companies": len(companies),
            "cost_centers": len(ccs), "levels": len(levels), "frequencies": len(freq)}


if __name__ == "__main__":
    out = None
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out") + 1]
    print(json.dumps(build(out), ensure_ascii=False))

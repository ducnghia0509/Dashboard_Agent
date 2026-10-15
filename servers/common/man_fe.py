# -*- coding: utf-8 -*-
"""Danh mục màn + NHÃN THẺ trên dashboard, đọc thẳng từ mã nguồn FE prod (15/10/2026).

VÌ SAO: bot kiểm số liệu từng bí ở câu "khối SRVF, màn Quản lý tồn kho, Giá trị tồn kho": (1) nó không
biết "Quản lý tồn kho" của Showroom là `vhkd4` nên hỏi ngược người dùng mã màn; (2) builder vhkd4 trả
số dưới tên trường (`tonHoaDonGiaTri`) còn nhãn "Giá trị tồn kho" chỉ nằm trong JSX của FE, nên dò từ
khoá trên JSON kết quả ra 0 mục. Module này đọc `nav.ts` (mã màn ↔ tên màn ↔ khối) và trang `.tsx`
(nhãn thẻ ↔ trường dữ liệu) để nối hai đầu đó. Đọc FE PROD (cây `apps/tc-console/tc-admin-web`) nên
nhãn luôn là nhãn người dùng đang thấy; không phải bảng ghim tay, FE đổi nhãn là module tự theo.
"""
import functools
import glob
import os
import re

from .mapping_check import _kd

WEB_SRC = os.environ.get("TC_WEB_SRC") or "/home/itadmin/apps/tc-console/tc-admin-web/src"

# Cách người dùng gọi khối -> chữ có trong tên nhóm màn của nav.ts.
_BI_DANH_KHOI = {
    "srvf": "vhkd", "showroom": "vhkd", "sr": "vhkd", "vinfast showroom": "vhkd", "ban xe": "vhkd",
    "xuong dich vu": "xdv", "xuong": "xdv", "xdv": "xdv",
    "an taxi": "an taxi", "antaxi": "an taxi", "xanh taxi": "xanh taxi", "xvp": "xanh taxi",
    "tram sac": "tram sac", "xe tai": "xe tai", "hung thinh": "xe tai", "du an": "du an",
    "qlts": "tai san", "tai san": "tai san",
}


def _doc(p):
    with open(p, encoding="utf-8") as fh:
        return fh.read()


@functools.lru_cache(maxsize=1)
def _nav(mtime):
    s = _doc(os.path.join(WEB_SRC, "nav.ts"))
    api = dict(re.findall(r"^\s*'?([\w-]+)'?:\s*'(\w+)',\s*$", s.split("export const MAN_API_KEY", 1)[1].split("};", 1)[0], re.M))
    khoi = dict(re.findall(r"^\s*'([\w-]+)':\s*'([^']+)',", s.split("export const KHOI_THEO_MAN", 1)[1].split("};", 1)[0], re.M))
    ra, nhom = [], ""
    for obj in re.findall(r"\{[^{}]*\}", s.split("export const NAV", 1)[1].split("];", 1)[0]):
        g = dict(re.findall(r"(\w+):\s*'([^']*)'", obj))
        id_, nhan, kind, path = g.get("id"), g.get("label"), g.get("kind"), g.get("path")
        if kind == "section":
            nhom = nhan
            continue
        if not path:
            continue                                  # nhóm có path (vd 'Hiệu quả kinh doanh' -> profit) VẪN là một màn
        man = api.get(path) or id_
        ra.append({"man": man, "ten_man": nhan, "nhom": nhom, "khoi": khoi.get(path), "duong_dan": path})
    return ra


def danh_muc_man():
    p = os.path.join(WEB_SRC, "nav.ts")
    return _nav(os.path.getmtime(p))


def tim_man(cau):
    """Tên/mã màn người dùng nói ("vhkd4", "Quản lý tồn kho SRVF", "tồn kho showroom") -> danh sách màn khớp,
    tốt nhất trước. Rỗng = không đoán được."""
    ds = danh_muc_man()
    # Người dùng dán URL/đường dẫn: "\\hieuqua?ky=ngay", ".../vhkd-tonkho" -> khớp thẳng theo path.
    for seg in reversed(re.split(r"[\\/]+", re.sub(r"[?#].*$", "", str(cau or "").strip()))):
        hit = [m for m in ds if seg and m["duong_dan"] == seg.lower()]
        if hit:
            return hit
    q = _kd(cau)
    exact = [m for m in ds if _kd(m["man"]) == q]
    if exact:
        return exact
    tu_khoi = {v for k, v in _BI_DANH_KHOI.items() if re.search(rf"\b{re.escape(k)}\b", q)}
    for k in _BI_DANH_KHOI:
        q = re.sub(rf"\b{re.escape(k)}\b", " ", q)
    q = re.sub(r"\b(man|khoi|tab|dashboard)\b", " ", q).split()

    def diem(m):
        ten, ctx = _kd(m["ten_man"]), _kd(f"{m['nhom']} {m['khoi'] or ''} {m['man']}")
        tu = ten.split()
        d = sum(1 for w in q if w in tu) / max(len(q), 1) - 0.01 * max(0, len(tu) - len(q))
        if tu_khoi:
            d += 0.5 if any(k in ctx or k in _kd(m["man"]) for k in tu_khoi) else -1
        return d
    xep = sorted(((diem(m), m) for m in ds), key=lambda x: -x[0])
    return [m for d, m in xep if d > 0.4][:5]


def _file_trang(man):
    for p in glob.glob(os.path.join(WEB_SRC, "pages", "*.tsx")):
        if p.endswith(".test.tsx"):
            continue
        if re.search(rf"usePageData<[^>]*>\(\s*'{re.escape(man)}'", _doc(p)):
            return p
    return None


@functools.lru_cache(maxsize=64)
def _nhan(p, mtime):
    s = _doc(p)
    ra, thay = [], set()
    # <Kpi label="Giá trị tồn kho" value={fmtTy(d.tonHoaDonGiaTri)} ...> — lấy trường `d.xxx` ĐẦU TIÊN ngay
    # trong thuộc tính value/các thuộc tính của chính thẻ đó (không vượt quá dấu '>' đóng thẻ).
    for m in re.finditer(r"""label=(?:"([^"]+)"|\{\s*'([^']+)'\s*\})([^>]{0,400})""", s):
        nhan = m.group(1) or m.group(2)
        f = re.search(r"\bd\??\.([A-Za-z_][\w]*(?:\??\.[A-Za-z_]\w*)*)", m.group(3))
        if not f:
            continue
        truong = f.group(1).replace("?.", ".")
        if (nhan, truong) not in thay:
            thay.add((nhan, truong))
            ra.append({"nhan": nhan, "truong": truong})
    return ra


def nhan_tren_man(man):
    """[{nhan, truong}] — nhãn thẻ trên màn và trường dữ liệu (tương đối với khối dữ liệu `d` của trang) nó đọc."""
    p = _file_trang(man)
    return _nhan(p, os.path.getmtime(p)) if p else []

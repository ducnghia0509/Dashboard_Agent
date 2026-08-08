# -*- coding: utf-8 -*-
"""Lớp adapter cho brief Chủ tịch: chạy LẠI các extractor/deriver có sẵn ở CHẾ ĐỘ OFFLINE
(không DB, không ghi file template) và thu lấy bản ghi đã chuẩn hoá.

Vì sao cần lớp này: 17 extractor trong scripts/ đã xử lý xong toàn bộ khác biệt layout giữa
các đơn vị (T-series HT, A-series SRVF, tcode GA, B-series XDV, layout Dự án...) — công sức
nhiều tháng. Viết lại logic đọc Excel cho brief là vừa trùng lặp vừa chắc chắn lệch số so với
dashboard. Nhưng mọi extractor đều kết thúc bằng `tf.fill(...)` rồi `tf.import_filled(...)`
để nạp DB, mà brief thì KHÔNG dùng DB (quyết định 07/08: đợt này agent chỉ đọc Excel).

Điểm chặn: `tf.fill(template, recs, out)` — đúng chỗ dữ liệu đã parse xong và đã chuẩn hoá về
template chuẩn, ngay TRƯỚC khi chạm đĩa/DB. Patch 3 điểm (tf.fill, tf.import_filled, và
be_bridge.db.get_db cho vài extractor có direct-insert) là lấy được toàn bộ bản ghi mà KHÔNG
sửa một dòng nào của pipeline production.
"""
import contextlib
import glob
import os
import re
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
# Vài deriver import ANH EM CÙNG THƯ MỤC theo tên trơn (`import agent_cli`) vì vốn được chạy
# dạng script `python scripts/x.py`. Import chúng dạng package (`from scripts import x`) sẽ vỡ
# nếu scripts/ không nằm trên sys.path.
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from servers.common import source_catalog as SC  # noqa: E402

RECEIVED_DIR = SC.RECEIVED_DIR


class _NullCursor:
    """Con trỏ DB rỗng — nuốt mọi execute/commit của các khối direct-insert trong extractor
    (vd extract_sodu_tien ghi thẳng DACHI_CCT). Trả rỗng thay vì ném lỗi để extractor chạy hết
    phần parse; brief chỉ cần recs từ tf.fill, không cần phần ghi DB đó."""

    def execute(self, *a, **k):
        return self

    def fetchone(self):
        return None

    def fetchall(self):
        return []

    def commit(self):
        pass

    def rollback(self):
        pass

    def close(self):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


@contextlib.contextmanager
def offline():
    """Tắt mọi đường ghi (template xlsx + DB) của extractor, thu lại recs từ tf.fill.

    Yield ra list `captured` gồm các tuple (template_name, recs). Extractor nào gọi tf.fill
    nhiều lần (nhiều template) thì có nhiều phần tử, giữ nguyên thứ tự gọi.
    """
    from servers import template_filler as tf
    from servers.common import be_bridge as bb

    captured = []
    orig_fill, orig_import = tf.fill, tf.import_filled
    orig_get_db = bb.db.get_db

    def _fill(template, recs, out=None, *a, **k):
        captured.append((template, list(recs or [])))
        return out

    tf.fill = _fill
    tf.import_filled = lambda *a, **k: {"rows_imported": 0, "by_type": {}, "offline": True}
    bb.db.get_db = lambda *a, **k: _NullCursor()
    try:
        yield captured
    finally:
        tf.fill, tf.import_filled = orig_fill, orig_import
        bb.db.get_db = orig_get_db


def capture(func, *args, **kwargs):
    """Chạy 1 extractor ở chế độ offline. Trả (result, recs_by_template).

    recs_by_template: dict {template_name: [rec, ...]}. Extractor lỗi -> ném lên cho người gọi
    quyết định (build_chutich_brief bắt và ghi vào _META.errors, KHÔNG nuốt im lặng).
    """
    with offline() as captured:
        result = func(*args, **kwargs)
    by_tpl = {}
    for tpl, recs in captured:
        by_tpl.setdefault(tpl, []).extend(recs)
    return result, by_tpl


# ---- Dò file nguồn ----------------------------------------------------------------

# Kỳ trong tên file: '.M.202607.' / '.D.202608.' / 'M202601' (thiếu dấu chấm, SRVF T01)
# / '.Y.2026.' (kế hoạch năm). Bắt theo THỨ TỰ ưu tiên tháng trước năm.
_RE_PERIOD = [
    (re.compile(r"[.\s]([MDY])[.\s]?(\d{6})(?:\d\d)?[.\s]"), "ym"),
    # Khối Dự án đặt tên tháng KHÔNG zero-pad: 'B.4.TC.TCKT.D.20268.' = tháng 8/2026, và
    # '...M.20267.' = tháng 7. Không bắt trường hợp này thì mọi file Dự án đều mất kỳ.
    (re.compile(r"[.\s]([MDY])[.\s]?(\d{4})([1-9])[.\s]"), "ym5"),
    (re.compile(r"[.\s]Y[.\s](\d{4})[.\s]"), "y"),
    (re.compile(r"THÁNG\s*(\d{1,2})\s*NĂM\s*(\d{4})", re.IGNORECASE), "vn"),
]


def parse_period(file_name: str):
    """-> (period 'YYYY-MM' hoặc 'YYYY', grain 'M'|'D'|'Y') hoặc (None, None).

    Lưu ý An Taxi dùng '.M.' cho báo cáo NGÀY (đã ghi nhận trong pipeline) nên grain suy từ tên
    file KHÔNG tuyệt đối tin được — người gọi nên lọc thêm bằng thư mục report_type.
    """
    for rx, kind in _RE_PERIOD:
        m = rx.search(file_name)
        if not m:
            continue
        if kind == "ym":
            grain, ym = m.group(1), m.group(2)
            return f"{ym[:4]}-{ym[4:6]}", grain
        if kind == "ym5":
            grain, y, mo = m.group(1), m.group(2), m.group(3)
            return f"{y}-{int(mo):02d}", grain
        if kind == "y":
            return m.group(1), "Y"
        if kind == "vn":
            return f"{m.group(2)}-{int(m.group(1)):02d}", "M"
    return None, None


def discover(report_type: str = None, folder: str = None) -> list:
    """Liệt kê file nguồn thật trong received_reports (KHÔNG mở workbook — chỉ đọc tên/mtime).

    Trả list dict {path, folder, report_type, file, period, grain, mtime} sắp theo period rồi
    mtime giảm dần (mới nhất trước).
    """
    pat = os.path.join(RECEIVED_DIR, folder or "*", report_type or "*", "*.xlsx")
    out = []
    for p in glob.glob(pat):
        base = os.path.basename(p)
        if base.startswith("~$"):          # file tạm của Excel
            continue
        rel = os.path.relpath(p, RECEIVED_DIR).split(os.sep)
        period, grain = parse_period(base)
        out.append({
            "path": p,
            "folder": rel[0] if len(rel) >= 2 else None,
            "report_type": rel[1] if len(rel) >= 3 else None,
            "file": base,
            "period": period,
            "grain": grain,
            "mtime": os.path.getmtime(p),
        })
    out.sort(key=lambda e: (e["period"] or "", e["mtime"]), reverse=True)
    return out


def latest_per_folder(report_type: str, period: str = None) -> dict:
    """-> {folder: entry} lấy file mới nhất của mỗi thư mục nguồn cho 1 report_type.

    Có `period` thì chỉ xét đúng kỳ đó; không thì lấy kỳ mới nhất mà từng thư mục có (các đơn vị
    nộp lệch nhau — ép chung 1 kỳ sẽ làm rỗng đơn vị nộp chậm).
    """
    best = {}
    for e in discover(report_type=report_type):
        if period and e["period"] != period:
            continue
        cur = best.get(e["folder"])
        if cur is None or (e["period"] or "", e["mtime"]) > (cur["period"] or "", cur["mtime"]):
            best[e["folder"]] = e
    return best


def latest_period(report_type: str) -> str:
    """Kỳ mới nhất đang có dữ liệu của 1 report_type (bỏ qua file không suy được kỳ)."""
    ps = [e["period"] for e in discover(report_type=report_type) if e["period"] and len(e["period"]) == 7]
    return max(ps) if ps else None


def prev_period(period: str) -> str:
    """'2026-07' -> '2026-06'. Dùng cho mọi phép so kỳ trước trong brief."""
    y, m = int(period[:4]), int(period[5:7])
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"

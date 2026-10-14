# -*- coding: utf-8 -*-
"""Xuất LOGIC lên số của từng màn dashboard ra JSON — chạy bằng `.venv` của tc-admin-api (cần import `app.*`).

Dùng cho bot so khớp mapping (`servers/common/mapping_check.py` gọi file này bằng subprocess rồi cache kết quả).
KHÔNG đọc DB. Mỗi màn: hàm build thật đứng sau `registry.SCREENS[man]`, file:dòng, docstring, các report_type
builder đọc (tĩnh, qua AST — đúng hàm `screen_sources.report_types_cua_man` mà thanh lọc đang dùng),
nguồn NGÀY riêng/chung, và đoạn mã nguồn đầu hàm (tối đa MAX_DONG dòng) để người/LLM đọc công thức.

    <tc-admin-api>/.venv/bin/python scripts/mapping_logic_dump.py > logic.json
"""
import inspect
import json
import os
import sys

API = os.environ.get("TC_API_DIR") or "/home/itadmin/apps/tc-console/tc-admin-api"
sys.path.insert(0, API)
os.environ.setdefault("DATABASE_URL", "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard")

MAX_DONG = 160


def main():
    from app.metrics import registry, screen_sources as ss
    ra = {"api_dir": API, "man": {}}
    for man in sorted(registry.SCREENS):
        fn = ss._builder(man)
        m = {"builder": None}
        if fn is not None:
            try:
                nguon, dong = inspect.getsourcelines(fn)
                m = {"builder": fn.__name__, "module": fn.__module__,
                     "file": os.path.relpath(inspect.getsourcefile(fn), API), "dong_bat_dau": dong,
                     "so_dong_ham": len(nguon), "doc": (inspect.getdoc(fn) or "")[:1800],
                     "ma_nguon": "".join(nguon[:MAX_DONG])}
            except (OSError, TypeError):
                m = {"builder": fn.__name__}
        rt = ss.report_types_cua_man(man)
        m["report_types"] = sorted(rt) if rt else []
        try:
            nd = ss.report_types_ngay_cua_man(man)
            m["rt_ngay"] = ({k: sorted(v) for k, v in nd.items()} if isinstance(nd, dict) else None)
        except Exception:                                   # noqa: BLE001
            m["rt_ngay"] = None
        ra["man"][man] = m
    print(json.dumps(ra, ensure_ascii=False))


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Ket noi Postgres GHI cho luong ingest (import_execute). Dung chung PgShim voi
db_ro.py, chi khac bien env - DATABASE_URL (quyen ghi, giong DashBoard_AI/backend/.env).

Uu tien: goi thang be_bridge.get_db() (dung nguyen connection cua DashBoard_AI qua
app/db.py) cho moi thao tac ghi raw_rows/datasets, de dam bao insert dung schema/
idempotency da co san (commit()/import_workbook()/repo.py). db_rw.py o day chi dung
cho cac truy van PHU (vd kiem tra dataset da ton tai truoc khi ghi) khi khong tien
goi thang qua be_bridge.
"""
from .db_ro import PgShim

_rw_singleton = None


def get_rw_db() -> PgShim:
    global _rw_singleton
    if _rw_singleton is None:
        _rw_singleton = PgShim("DATABASE_URL", "luồng ingest (ghi raw_rows)")
    return _rw_singleton

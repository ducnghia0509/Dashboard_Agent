# -*- coding: utf-8 -*-
"""Ket noi Postgres READ-ONLY cho luong qa (sql_query). Rut gon tu
DashBoard_AI/backend/app/db.py: psycopg3 + dict_row + shim `?`->`%s`.

DUNG DATABASE_URL_RO - PHAI la 1 role da GRANT SELECT-only (xem sql/create_ro_role.sql).
KHONG fallback sang DATABASE_URL neu thieu bien nay - fail loud de khong vo tinh
dung nham quyen ghi cho luong hoi dap.
"""
import os
import threading

import psycopg
from psycopg.rows import dict_row


class PgShim:
    """Wrapper mong: dich `?`->`%s`, tu ket noi lai khi rot, API kieu sqlite3.Connection."""

    def __init__(self, env_var: str, purpose: str):
        self.env_var = env_var
        self.purpose = purpose
        self._local = threading.local()

    @property
    def _dsn(self):
        dsn = os.environ.get(self.env_var)
        if not dsn:
            raise RuntimeError(
                f"Chưa cấu hình {self.env_var} (Postgres, dùng cho {self.purpose}). "
                f"Copy .env.example → .env và điền connection string."
            )
        return dsn

    def _connect(self) -> psycopg.Connection:
        return psycopg.connect(
            self._dsn, autocommit=True, prepare_threshold=None, row_factory=dict_row,
        )

    @property
    def raw(self) -> psycopg.Connection:
        conn = getattr(self._local, "conn", None)
        if conn is None or conn.closed:
            conn = self._connect()
            self._local.conn = conn
        return conn

    @staticmethod
    def _sql(q: str) -> str:
        return q.replace("?", "%s")

    def execute(self, q: str, params=None):
        sql = self._sql(q)
        try:
            cur = self.raw.cursor()
            cur.execute(sql, tuple(params) if params is not None else None)
            return cur
        except (psycopg.OperationalError, psycopg.InterfaceError):
            self._local.conn = None
            cur = self.raw.cursor()
            cur.execute(sql, tuple(params) if params is not None else None)
            return cur


_ro_singleton = None


def get_ro_db() -> PgShim:
    global _ro_singleton
    if _ro_singleton is None:
        _ro_singleton = PgShim("DATABASE_URL_RO", "luồng qa (đọc read-only)")
    return _ro_singleton

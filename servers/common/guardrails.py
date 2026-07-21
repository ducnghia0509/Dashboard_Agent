# -*- coding: utf-8 -*-
"""Guardrails cho sql_query (luong qa): chi cho SELECT/WITH read-only, chan DML/DDL,
chan multi-statement, ep LIMIT, ap statement_timeout. Day la lop phong thu THU HAI
(lop 1 la role Postgres read-only - xem sql/create_ro_role.sql)."""
import os
import re

_BLACKLIST = re.compile(
    r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|TRUNCATE|GRANT|REVOKE|COPY|"
    r"VACUUM|CALL|DO|EXECUTE|LISTEN|NOTIFY|SET|RESET|LOCK|"
    # INTO: 'SELECT ... INTO bang_moi' TẠO BẢNG dù câu bắt đầu bằng SELECT (lọt check
    # leading-keyword). FOR UPDATE/SHARE khoá dòng. Các lệnh ghi khác (MERGE/DECLARE/...)
    # đứng đầu câu nên đã bị _LEADING_KEYWORD_RE chặn, không cần thêm (tránh false-positive
    # kiểu 'FETCH FIRST 10 ROWS ONLY' là SELECT hợp lệ).
    r"INTO)\b"
    r"|\bFOR\s+(UPDATE|SHARE)\b",
    re.IGNORECASE,
)
_LIMIT_RE = re.compile(r"\bLIMIT\s+(\d+)\b", re.IGNORECASE)
_LEADING_KEYWORD_RE = re.compile(r"^\s*(SELECT|WITH)\b", re.IGNORECASE)
_COMMENT_RE = re.compile(r"--[^\n]*|/\*.*?\*/", re.DOTALL)


class GuardrailError(ValueError):
    pass


def _strip_comments(sql: str) -> str:
    return _COMMENT_RE.sub(" ", sql)


def check_select_only(sql: str) -> str:
    """Trả về SQL đã strip comment + trailing ';' nếu hợp lệ, raise GuardrailError nếu không."""
    if not sql or not sql.strip():
        raise GuardrailError("SQL rỗng.")
    clean = _strip_comments(sql).strip()
    # Chỉ cho phép đúng 1 dấu ';' ở cuối câu (nếu có) - chặn multi-statement.
    body = clean[:-1].rstrip() if clean.endswith(";") else clean
    if ";" in body:
        raise GuardrailError("Không cho phép nhiều câu lệnh (dấu ';' giữa câu).")
    if not _LEADING_KEYWORD_RE.match(body):
        raise GuardrailError("Chỉ cho phép câu lệnh bắt đầu bằng SELECT hoặc WITH.")
    if _BLACKLIST.search(body):
        raise GuardrailError("Câu lệnh chứa từ khoá không được phép (DML/DDL/session control).")
    return body


def ensure_limit(sql: str, default: int = None, max_limit: int = None) -> str:
    default = default or int(os.environ.get("SQL_DEFAULT_LIMIT", 500))
    max_limit = max_limit or int(os.environ.get("SQL_MAX_LIMIT", 2000))
    matches = list(_LIMIT_RE.finditer(sql))
    if not matches:
        return f"{sql.rstrip()} LIMIT {default}"
    # Chỉ xét LIMIT CUỐI CÙNG: LIMIT của câu ngoài (nếu có) luôn đứng sau LIMIT trong
    # subquery. Trước đây search() lấy LIMIT ĐẦU TIÊN -> 'WHERE id IN (SELECT ... LIMIT 5)'
    # bị coi là "đã có LIMIT" và câu ngoài chạy KHÔNG giới hạn; còn sub() không count thì
    # kẹp nhầm cả LIMIT của subquery.
    last = matches[-1]
    tail = sql[last.end():]
    if ")" in tail:
        # LIMIT cuối vẫn nằm TRONG ngoặc (subquery) -> câu ngoài chưa có LIMIT.
        return f"{sql.rstrip()} LIMIT {default}"
    n = int(last.group(1))
    if n > max_limit:
        return sql[:last.start()] + f"LIMIT {max_limit}" + tail
    return sql


def sanitize(sql: str, default_limit: int = None, max_limit: int = None) -> str:
    """check_select_only + ensure_limit, dùng 1 lần trước khi execute."""
    return ensure_limit(check_select_only(sql), default_limit, max_limit)


def run_readonly(conn, sql: str, params=None, timeout_ms: int = None):
    """Thực thi SQL đã sanitize trên connection read-only, có statement_timeout."""
    timeout_ms = timeout_ms or int(os.environ.get("SQL_STATEMENT_TIMEOUT_MS", 5000))
    safe_sql = sanitize(sql)
    conn.execute(f"SET statement_timeout = {int(timeout_ms)}")
    cur = conn.execute(safe_sql, params)
    rows = [dict(r) for r in cur.fetchall()]
    return rows, safe_sql

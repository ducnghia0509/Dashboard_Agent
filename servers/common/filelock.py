# -*- coding: utf-8 -*-
"""KHOÁ NGĂN 2 LƯỢT INGEST CÙNG 1 FILE CHỒNG THỜI GIAN NHAU (2026-08-07).

Sự cố gốc: `template_filler.import_filled()` (dùng bởi MỌI đường ingest tất định — web
/source/analyze|reprocess, cron_thuchi_daily.py, cron_hqkdngay_daily.py, agent_cli chạy tay)
ghi đè ĐÚNG cho ingest TUẦN TỰ: mỗi lượt tự chụp `before_id = MAX(id)`, ghi dòng mới, rồi xoá
dòng cũ có `id<=before_id` (xoá SAU khi ghi — cố ý, tránh khoảng trống nếu ghi lỗi giữa chừng).
Nhưng nếu 2 lượt ingest CÙNG 1 FILE chồng thời gian (bấm 2 lần liên tiếp, hoặc cron chạy trùng
lúc user bấm tay), mỗi lượt chụp `before_id` riêng TRƯỚC khi lượt kia ghi xong -> không lượt nào
xoá được dòng mới của lượt kia -> CẢ 2 BỘ DÒNG CÙNG TỒN TẠI, dữ liệu nhân đôi thật trong raw_rows
(không có ràng buộc UNIQUE nào chặn). Đây là hồi quy: bản backend cũ (DashBoard_AI/backend) từng
có `_ANALYZE_LOCK` chặn đúng việc này nhưng bị rớt khi tách sang tc-admin-api.

DÙNG flock (không dùng Postgres advisory lock): tầng kết nối DB dùng autocommit=True qua
Neon/PgBouncer, có cơ chế tự âm thầm reconnect khi rớt kết nối (xem be_bridge/db_rw) — khiến
session-level advisory lock có thể "bốc hơi" giữa chừng mà KHÔNG báo lỗi, còn nguy hiểm hơn race
đang chặn. flock cấp file (cùng cách `memory.locked_json()` đã dùng cho JSON dùng chung) an toàn
hơn ở đây và TỰ NHẢ khi tiến trình giữ khoá bị kill/crash (kể cả SIGKILL do subprocess timeout)
- không để lại khoá chết như cách dùng file PID.

CHỜ CÓ GIỚI HẠN, không fail ngay: ingest lại đúng file vài giây sau (kế toán sửa file, upload
lại) là tình huống HỢP LỆ, không phải lỗi — nên khoá poll tới khi hết `timeout` rồi mới báo lỗi,
không từ chối ngay lập tức."""
import fcntl
import hashlib
import os
import time
from contextlib import contextmanager

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
_LOCK_DIR = os.path.join(_AGENT_ROOT, ".ingest_locks")
_POLL_SEC = 0.5


class IngestInProgress(Exception):
    """Không xin được khoá trong thời gian chờ — có lượt ingest KHÁC đang xử lý CÙNG file này."""


def _lock_path(source_id: str) -> str:
    slug = hashlib.sha1(source_id.encode("utf-8")).hexdigest()
    return os.path.join(_LOCK_DIR, f"{slug}.lock")


@contextmanager
def ingest_lock(source_id: str, timeout: float = 90.0):
    """Giữ khoá EXCLUSIVE riêng cho `source_id` (vd '<công_ty_thư_mục>::<tên_file>') trong suốt
    khối `with` — 2 lượt ingest CÙNG source_id sẽ tự xếp hàng; source_id KHÁC nhau không đụng
    nhau (không khoá nguyên cả batch reprocess-nhiều-file). Raise IngestInProgress nếu hết
    `timeout` giây vẫn không xin được khoá (lượt kia treo/quá lâu)."""
    os.makedirs(_LOCK_DIR, exist_ok=True)
    path = _lock_path(source_id)
    fd = os.open(path, os.O_CREAT | os.O_RDWR, 0o644)
    deadline = time.time() + timeout
    try:
        while True:
            try:
                fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if time.time() >= deadline:
                    raise IngestInProgress(
                        f"'{source_id}' đang được xử lý bởi 1 lượt ingest khác "
                        f"(chờ {timeout:.0f}s không được) — thử lại sau."
                    )
                time.sleep(_POLL_SEC)
        try:
            yield
        finally:
            fcntl.flock(fd, fcntl.LOCK_UN)
    finally:
        os.close(fd)

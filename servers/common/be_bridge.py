# -*- coding: utf-8 -*-
"""Cau noi sang backend/app cua DashBoard_AI: sys.path.insert roi import lai nguyen
logic detect/auto_map/prepare/commit/import_workbook/master_data.

KHONG viet lai logic nghiep vu - moi thu goi thang vao module app.* cua DashBoard_AI.
"""
import os
import sys

from dotenv import load_dotenv

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
load_dotenv(os.path.join(_AGENT_ROOT, ".env"))

BACKEND_PATH = os.environ.get("BACKEND_PATH") or os.path.normpath(
    os.path.join(_AGENT_ROOT, "..", "DashBoard_AI", "backend")
)
BACKEND_PATH = os.path.normpath(os.path.join(_AGENT_ROOT, BACKEND_PATH)) \
    if not os.path.isabs(BACKEND_PATH) else BACKEND_PATH

if BACKEND_PATH not in sys.path:
    sys.path.insert(0, BACKEND_PATH)

try:
    from app import schemas, importer, importer_month, importer_ledger, db, repo, master, metrics, text_util
except ModuleNotFoundError as exc:  # pragma: no cover - lỗi cấu hình, không phải logic
    raise RuntimeError(
        f"Không import được app.* từ BACKEND_PATH={BACKEND_PATH!r}. "
        f"Kiểm tra .env (BACKEND_PATH) hoặc chạy `pip install -r {BACKEND_PATH}/../requirements.txt`."
    ) from exc

# Re-export ngắn gọn cho 2 MCP server dùng.
detect = schemas.detect
auto_map = schemas.auto_map
build_schemas = schemas.build_schemas
FIELD_DEFS = schemas.FIELD_DEFS
FIELD_LABELS = schemas.FIELD_LABELS
REPORT_CODES = schemas.REPORT_CODES
REPORT_LABELS = schemas.REPORT_LABELS
SNAPSHOT = schemas.SNAPSHOT

prepare = importer.prepare
commit = importer.commit
import_workbook = importer_month.import_workbook
detect_ledger = importer_ledger.detect_ledger

master_data = master.master_data
report_templates = master.report_templates
khoi_names = master.khoi_names
cc_to_khoi = master.cc_to_khoi

get_db = db.get_db
DB_SCHEMA = db.SCHEMA

normalize_header = text_util.normalize_header
remove_diacritics = text_util.remove_diacritics
parse_num = text_util.parse_num
parse_date = text_util.parse_date
parse_text = text_util.parse_text

fingerprint = repo.fingerprint
find_profile = repo.find_profile
save_profile = repo.save_profile
get_dataset = repo.get_dataset
list_datasets = repo.list_datasets
new_dataset = repo.new_dataset

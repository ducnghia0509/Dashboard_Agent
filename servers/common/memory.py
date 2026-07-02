# -*- coding: utf-8 -*-
"""DISCOVERY MEMORY (artifact A trong plan.md) - luu duoi dang JSON, KHONG vector,
tra cuu chinh xac theo ten file / report_type / tu khoa.

discovery_record (ghi, dung trong ingest_server) + discovery_search (doc, dung
trong qa_server) deu di qua module nay."""
import hashlib
import json
import os
from datetime import datetime, timezone

from . import be_bridge as bb

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
DISCOVERIES_DIR = os.path.join(_AGENT_ROOT, "memory", "discoveries")
REPORT_SPECS_DIR = os.path.join(_AGENT_ROOT, "memory", "report_specs")
BINDINGS_DIR = os.path.join(_AGENT_ROOT, "memory", "bindings")


def _slug(file_name: str) -> str:
    return hashlib.sha1(file_name.encode("utf-8")).hexdigest()[:16]


def _path_for(file_name: str) -> str:
    os.makedirs(DISCOVERIES_DIR, exist_ok=True)
    return os.path.join(DISCOVERIES_DIR, f"{_slug(file_name)}.json")


def discovery_record(
    file_name: str,
    fingerprint: str,
    sheets: list,
    columns_per_sheet: dict,
    detected_report_type: str = None,
    header_row: int = None,
    mapping: dict = None,
    period: str = None,
    confidence: float = 0.0,
    anomalies: list = None,
) -> dict:
    record = {
        "file_name": file_name,
        "fingerprint": fingerprint,
        "sheets": sheets,
        "columns_per_sheet": columns_per_sheet,
        "detected_report_type": detected_report_type,
        "header_row": header_row,
        "mapping": mapping or {},
        "period": period,
        "confidence": confidence,
        "anomalies": anomalies or [],
        "created_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    with open(_path_for(file_name), "w", encoding="utf-8") as fh:
        json.dump(record, fh, ensure_ascii=False, indent=2)
    return record


def _all_records() -> list:
    if not os.path.isdir(DISCOVERIES_DIR):
        return []
    out = []
    for fn in sorted(os.listdir(DISCOVERIES_DIR)):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(DISCOVERIES_DIR, fn), encoding="utf-8") as fh:
                out.append(json.load(fh))
        except (OSError, json.JSONDecodeError):
            continue
    return out


def discovery_search(query: str = None, report_type: str = None) -> list:
    """Tìm theo tên file (substring, không dấu/không phân biệt hoa thường) hoặc report_type."""
    records = _all_records()
    if report_type:
        records = [r for r in records if r.get("detected_report_type") == report_type]
    if query:
        nq = _norm(query)
        records = [r for r in records if nq in _norm(r.get("file_name", ""))
                   or nq in _norm(json.dumps(r, ensure_ascii=False))]
    return records


def _norm(s: str) -> str:
    return bb.normalize_header(s, True)


# ---- REPORT SPEC CATALOG (Extension 2) — catalog tích luỹ các SheetMapping đã học ----
# Khác discovery memory (ghi lại "đã thấy file gì"), catalog này ghi lại "cách lấy dữ liệu"
# (mapping khai báo) cho 1 sheet lạ đã được analyst suy luận + execute ghi thành công,
# để lần sau gặp sheet cùng fingerprint không phải suy luận lại từ đầu.

def _report_spec_path(fingerprint: str) -> str:
    os.makedirs(REPORT_SPECS_DIR, exist_ok=True)
    return os.path.join(REPORT_SPECS_DIR, f"{fingerprint[:16]}.json")


def report_spec_save(fingerprint: str, sheet_mapping: dict, verified: bool = True) -> dict:
    """Lưu 1 SheetMapping đã học vào catalog. verified=True nghĩa là mapping này đã được
    dùng để GHI THẬT vào raw_rows (đáng tin cậy nhất). verified=False nghĩa là mapping mới
    chỉ qua dry-run thành công (row_count>0) — khép vòng học sớm hơn (không phải chờ người
    approve ghi thật mới học được), nhưng CHƯA chắc chắn 100% nên autobatch vẫn có thể muốn
    ưu tiên bản verified=True nếu có nhiều bản ghi trùng canonical_kind/sheet."""
    existing = None
    path = _report_spec_path(fingerprint)
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as fh:
                existing = json.load(fh)
        except (OSError, json.JSONDecodeError):
            existing = None
    # Không để 1 dry-run (verified=False) đè lên 1 bản ghi đã verified=True trước đó.
    if existing and existing.get("verified") and not verified:
        return existing
    record = {
        "fingerprint": fingerprint, "sheet_mapping": sheet_mapping, "verified": verified,
        "saved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(record, fh, ensure_ascii=False, indent=2)
    return record


def _all_report_specs() -> list:
    if not os.path.isdir(REPORT_SPECS_DIR):
        return []
    out = []
    for fn in sorted(os.listdir(REPORT_SPECS_DIR)):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(REPORT_SPECS_DIR, fn), encoding="utf-8") as fh:
                out.append(json.load(fh))
        except (OSError, json.JSONDecodeError):
            continue
    return out


def report_spec_search(query: str = None, sheet: str = None, target_report_type: str = None,
                        canonical_kind: str = None) -> list:
    """Tìm mapping đã học theo tên sheet / report_type / canonical_kind / từ khoá tự do
    (không dấu). canonical_kind (vd 'TK131') tìm được BẤT KỂ tên sheet/file/công ty khác
    nhau — dùng khi công ty khác đặt tên sheet khác nhưng cùng loại báo cáo kế toán."""
    records = _all_report_specs()
    if sheet:
        records = [r for r in records if r["sheet_mapping"].get("sheet") == sheet]
    if target_report_type:
        records = [r for r in records
                   if r["sheet_mapping"].get("target_report_type") == target_report_type]
    if canonical_kind:
        records = [r for r in records if r["sheet_mapping"].get("canonical_kind") == canonical_kind]
    if query:
        nq = _norm(query)
        records = [r for r in records if nq in _norm(json.dumps(r, ensure_ascii=False))]
    return records


# ---- Extension 3: bindings — xác nhận scope/basis/screen cho 1 canonical_kind ----
# Khác report_specs (mapping cột→raw_rows, kỹ thuật): binding là QUYẾT ĐỊNH NGHIỆP VỤ do
# NGƯỜI xác nhận 1 lần (hợp nhất/riêng, lũy kế/theo kỳ, phục vụ chỉ tiêu/màn hình nào) — lưu
# theo canonical_kind (không theo fingerprint file) để dùng lại cho MỌI công ty/kỳ cùng loại
# báo cáo, không phải hỏi lại. KHÔNG tự động ghi số vào KPI — chỉ là siêu dữ liệu phân loại.

def _binding_path(canonical_kind: str) -> str:
    os.makedirs(BINDINGS_DIR, exist_ok=True)
    return os.path.join(BINDINGS_DIR, f"{canonical_kind}.json")


def binding_get(canonical_kind: str):
    path = _binding_path(canonical_kind)
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return None


def binding_save(canonical_kind: str, scope: str = None, basis: str = None,
                 target_screen: str = None, chi_tieu: str = None, kpi_id=None) -> dict:
    record = {
        "canonical_kind": canonical_kind, "scope": scope, "basis": basis,
        "target_screen": target_screen, "chi_tieu": chi_tieu, "kpi_id": kpi_id,
        "confirmed": True,
        "confirmed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    with open(_binding_path(canonical_kind), "w", encoding="utf-8") as fh:
        json.dump(record, fh, ensure_ascii=False, indent=2)
    return record

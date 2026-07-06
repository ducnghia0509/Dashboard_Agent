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


# ---- Unmapped cost center registry ----
# Khi điền template gặp cost center/bộ phận KHÔNG khớp MD_COSTCENTER: vẫn nạp số (tổng
# đúng, dòng gom vào "(Chưa phân bổ)") NHƯNG ghi lại đây để admin bổ sung danh mục;
# lần sau CC đó tự roll-up đúng khối. Dedup theo chuỗi gốc (đã chuẩn hoá).

_UNMAPPED_CC = os.path.join(_AGENT_ROOT, "memory", "unmapped_cc.json")


def _load_unmapped() -> list:
    if not os.path.exists(_UNMAPPED_CC):
        return []
    try:
        with open(_UNMAPPED_CC, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return []


def unmapped_cc_record(raw: str, sheet: str = None, file_name: str = None,
                       cong_ty: str = None) -> dict:
    """Ghi 1 cost center chưa map (dedup theo raw). Tăng count nếu đã có."""
    raw = (raw or "").strip()
    logs = _load_unmapped()
    key = bb.remove_diacritics(raw).lower()
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    for e in logs:
        if bb.remove_diacritics(e.get("raw", "")).lower() == key:
            e["count"] = e.get("count", 1) + 1
            e["last_seen"] = now
            if sheet and sheet not in e.setdefault("sheets", []):
                e["sheets"].append(sheet)
            break
    else:
        logs.append({"raw": raw, "cong_ty": cong_ty, "sheets": [sheet] if sheet else [],
                     "file_name": file_name, "count": 1,
                     "first_seen": now, "last_seen": now, "resolved": False})
    os.makedirs(os.path.dirname(_UNMAPPED_CC), exist_ok=True)
    with open(_UNMAPPED_CC, "w", encoding="utf-8") as fh:
        json.dump(logs, fh, ensure_ascii=False, indent=2)
    return {"raw": raw, "total_unmapped": sum(1 for e in logs if not e.get("resolved"))}


def unmapped_cc_list(include_resolved: bool = False) -> list:
    return [e for e in _load_unmapped() if include_resolved or not e.get("resolved")]


# ---- Fill specs — mapping nguồn->template ĐÃ HỌC (khép vòng P2<->P3) ----
# Khi analyst điền template thành công 1 layout, lưu mapping theo FINGERPRINT của sheet nguồn
# (tên sheet + header cột). Lần sau file CÙNG layout (kể cả công ty/kỳ khác) -> orchestrator tự
# fill+import KHÔNG cần LLM. Đây là đòn bẩy scale chính (analyst chỉ chạy cho layout MỚI).

FILL_SPECS_DIR = os.path.join(_AGENT_ROOT, "memory", "fill_specs")


def source_fingerprint(sheet_name: str, columns: list) -> str:
    """Fingerprint 1 layout sheet nguồn = sha1(tên sheet chuẩn hoá + các cột header chuẩn hoá)."""
    norm = lambda s: bb.remove_diacritics("" if s is None else str(s)).strip().lower()
    key = norm(sheet_name) + "|" + "|".join(norm(c) for c in columns if c not in (None, ""))
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]


def fill_spec_save(fingerprint: str, target_sheet: str, mapping: dict,
                   source_sheet: str = None, canonical_kind: str = None,
                   value_scale: float = 1.0, constants: dict = None,
                   rename_rows: dict = None) -> dict:
    os.makedirs(FILL_SPECS_DIR, exist_ok=True)
    rec = {"fingerprint": fingerprint, "source_sheet": source_sheet,
           "target_sheet": target_sheet, "mapping": mapping, "canonical_kind": canonical_kind,
           "value_scale": value_scale, "constants": constants or {},
           "rename_rows": rename_rows or {},   # đổi tên dòng->chuẩn (KQKD), replay khi autofill
           "saved_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
    with open(os.path.join(FILL_SPECS_DIR, f"{fingerprint}.json"), "w", encoding="utf-8") as fh:
        json.dump(rec, fh, ensure_ascii=False, indent=2)
    return rec


def fill_spec_find(fingerprint: str):
    path = os.path.join(FILL_SPECS_DIR, f"{fingerprint}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return None

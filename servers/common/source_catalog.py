# -*- coding: utf-8 -*-
"""SOURCE CATALOG — index (lossless pointer) mọi file xlsx đã kéo về (Connect_VPS/received_reports).

Bronze layer = FILE TRÊN ĐĨA (không nhân bản vào DB). Catalog chỉ lưu con trỏ + cấu trúc
(sheet/cột/số dòng/canonical_kind) để QA/analyst biết "có file/sheet/cột nào" tức thì, kể cả
file CHƯA import vào raw_rows. Chi tiết ô đọc on-demand bằng source_inspect.

Index dựng lúc file land (P3). Truy vấn qua catalog_search (qa_server).
"""
import glob
import hashlib
import json
import os
from datetime import datetime, timezone


from . import be_bridge as bb
from . import canonical
from . import contract
from .memory import atomic_dump_json, locked_json

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))

# Thư mục file kéo về (đổi tên Connect_VPN -> Connect_VPS). Override qua env RECEIVED_DIR.
RECEIVED_DIR = os.environ.get("RECEIVED_DIR") or os.path.normpath(
    os.path.join(_AGENT_ROOT, "..", "Connect_VPS", "received_reports"))
CATALOG = os.path.join(_AGENT_ROOT, "memory", "source_catalog.json")


def _norm(s) -> str:
    return bb.remove_diacritics("" if s is None else str(s)).strip().lower()


def _load() -> dict:
    if not os.path.exists(CATALOG):
        return {}
    try:
        with open(CATALOG, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {}


def _save(cat: dict):
    atomic_dump_json(cat, CATALOG)


def _file_key(path: str) -> str:
    return hashlib.sha1(os.path.abspath(path).encode("utf-8")).hexdigest()[:16]


def _sidecar(path: str) -> dict:
    """File .json cùng tên (do receiver ghi) chứa company/month/report_type."""
    j = os.path.splitext(path)[0] + ".json"
    if os.path.exists(j):
        try:
            with open(j, encoding="utf-8") as fh:
                return json.load(fh)
        except (OSError, json.JSONDecodeError):
            pass
    return {}


def _from_path(path: str) -> dict:
    """Suy company/report_type từ cấu trúc received_reports/{company}/{report_type}/file.

    B28: segment đầu KHÔNG LUÔN LÀ pháp nhân — có thể là folder phân loại báo cáo (vd 'THUCHI'
    cho báo cáo thu-chi hợp nhất Tập đoàn, nhiều pháp nhân/sheet trong 1 file). company_guess
    giữ RAW để debug/hiển thị; company = None nếu segment không khớp MD_CONGTY (companies.yaml)
    — KHÔNG để analyst/QA hiểu lầm 'THUCHI' là 1 pháp nhân thật (đã xảy ra, gây ghi sai cong_ty)."""
    rel = os.path.relpath(path, RECEIVED_DIR)
    parts = rel.split(os.sep)
    # B27: file đặt nông received_reports/{company}/file (2 phần) vẫn suy được company.
    raw_company = parts[0] if len(parts) >= 2 else None
    # B30: prefer_file_name=True — TÊN FOLDER không đáng tin (đã xác nhận: chỉ là tên khối/loại
    # báo cáo do sender đặt, có lúc lẫn nhiều công ty trong 1 folder — vd 'HO/' chứa cả GA lẫn TC).
    # Tên FILE theo quy ước 'B.<khối>.<mã cty>.' do nghiệp vụ đặt, đáng tin hơn — luôn ưu tiên.
    return {"company": contract.resolve_company(raw=raw_company, file_name=os.path.basename(path),
                                                  prefer_file_name=True),
            "company_guess_raw": raw_company,
            "report_type": parts[1] if len(parts) >= 3 else None}


def raw_company_from_path(path: str):
    """Token CÔNG TY RAW = tên thư mục ngay dưới received_reports/ (vd 'HTXXANHTUYENQUANG').
    KHÁC _from_path().company (đã resolve theo TÊN FILE về mã pháp nhân) — nhiều nguồn thật khác
    nhau share cùng mã trong tên file (vd 3 HTX đều nộp 'B.6.XVP...'). Đồng bộ với
    sync_orchestrator._raw_company_from_path (nguồn của source_key trên UI)."""
    try:
        rel = os.path.relpath(os.path.abspath(path), RECEIVED_DIR)
    except (ValueError, TypeError):
        return None
    if rel.startswith(".."):
        return None
    parts = rel.split(os.sep)
    return parts[0] if len(parts) >= 2 else None


def source_id_from_path(path: str) -> str:
    """ĐỊNH DANH NGUỒN DUY NHẤT dùng làm raw_rows.source_file = '<công_ty_thư_mục>::<tên_file>'
    khi file nằm trong received_reports/<công_ty>/... — KHỚP source_key mà sync_orchestrator sinh
    cho UI, để trạng thái/ẩn-hiện/idempotent-delete khớp ĐÚNG TỪNG NGUỒN. File ngoài
    received_reports (vd upload tay) -> trả tên file trơn (không có thư mục nguồn để phân biệt).

    Lý do: nhiều nguồn thật (HTXXANHTUYENQUANG / HTXXANHVINHPHUC / XANHVINHPHUC) nộp CÙNG tên
    'B.6.XVP...xlsx' nhưng nội dung KHÁC; nếu source_file chỉ là basename thì 3 nguồn đè/che nhau
    (idempotent-delete xoá nhầm, trạng thái & ẩn/hiện lẫn lộn)."""
    base = os.path.basename(path or "")
    raw = raw_company_from_path(path)
    return f"{raw}::{base}" if raw else base


def index_file(path: str) -> dict:
    """Index 1 file xlsx -> entry {file, path, company, report_type, month, sheets:[...], ...}."""
    side = _sidecar(path)
    meta = _from_path(path)
    # B29/B30: sidecar .json do RECEIVER ngoài ghi — đã xác nhận (2026-07-09) đây chỉ là COPY
    # THẲNG 1 template/sidecar mẫu cũ (field "company" có thể mang giá trị VALID NHƯNG SAI, vd
    # copy nguyên "GA" cho file TC, hoặc theo tên folder cha 'HO' — không phải công ty thật).
    # KHÔNG tin sidecar/folder — LUÔN ưu tiên tên FILE (quy ước 'B.<khối>.<mã cty>.', do nghiệp
    # vụ đặt tên, đáng tin nhất) qua prefer_file_name=True; sidecar chỉ dùng khi tên file không
    # suy được (hiếm).
    side_company = contract.resolve_company(raw=side.get("company"), file_name=os.path.basename(path),
                                              prefer_file_name=True)
    entry = {
        "file": os.path.basename(path),
        "path": os.path.abspath(path),
        "company": side_company or meta.get("company"),
        "report_type": side.get("report_type") or meta.get("report_type"),
        "month": side.get("month"),
        "period_type": side.get("period_type"),
        "sheets": [],
        "mtime": os.path.getmtime(path),
        "indexed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "ingested": False,
    }
    wb = bb.fast_load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            # CHỈ đọc ~30 dòng đầu để lấy header (KHÔNG duyệt hết — sheet BCTC có thể tới
            # cả triệu dòng phantom -> duyệt hết là treo). Số dòng lấy từ ws.max_row.
            header = []
            for r in ws.iter_rows(min_row=1, max_row=30, values_only=True):
                if sum(1 for c in r if c not in (None, "")) >= 2:
                    header = [("" if c is None else str(c).strip()) for c in r]
                    break
            entry["sheets"].append({
                "name": ws.title,
                "columns": [h for h in header if h][:40],
                "nrows": ws.max_row,          # xấp xỉ (dimension Excel), tránh duyệt toàn bộ
                "canonical_kind": canonical.guess_canonical_kind(ws.title),
            })
    finally:
        wb.close()
    # Lock cả chu trình load-sửa-save: 2 phiên index 2 file song song không khoá sẽ
    # lost-update (mỗi bên save catalog thiếu entry của bên kia).
    with locked_json(CATALOG):
        cat = _load()
        key = _file_key(path)
        if key in cat:                      # giữ cờ ingested nếu đã có
            entry["ingested"] = cat[key].get("ingested", False)
        cat[key] = entry
        _save(cat)
    return entry


def index_dir(root: str = None) -> dict:
    """Quét thư mục received_reports, index mọi .xlsx (bỏ file tạm ~$)."""
    root = root or RECEIVED_DIR
    if not os.path.isdir(root):
        return {"ok": False, "error": f"Không thấy thư mục: {root}", "indexed": 0}
    files = [f for f in glob.glob(os.path.join(root, "**", "*.xlsx"), recursive=True)
             if not os.path.basename(f).startswith("~$")]
    cat = _load()
    done = 0
    for f in files:
        try:
            key = _file_key(f)
            prev = cat.get(key)
            # BỎ QUA file đã index & chưa đổi (mtime khớp) -> tránh mở lại file lớn (17MB) mỗi lần.
            if prev and prev.get("mtime") == os.path.getmtime(f):
                continue
            index_file(f)
            done += 1
        except Exception:  # 1 file hỏng không chặn cả mẻ
            pass
    # B22: PRUNE entry trỏ file đã biến mất trên đĩa (tránh reconcile/status báo "đã nạp" sai,
    # QA mở path chết) — khớp triệu chứng dir ANTAXI/HO/TRAMSAC rỗng mà catalog còn trỏ.
    with locked_json(CATALOG):
        cat = _load()
        removed = [k for k, e in cat.items() if not os.path.exists(e.get("path", ""))]
        for k in removed:
            del cat[k]
        if removed:
            _save(cat)
    return {"ok": True, "scanned": len(files), "indexed_new": done,
            "pruned": len(removed), "total_in_catalog": len(cat)}


def mark_ingested(path: str, ingested: bool = True):
    with locked_json(CATALOG):
        cat = _load()
        key = _file_key(path)
        if key in cat:
            cat[key]["ingested"] = ingested
            _save(cat)


def search(query: str = None, company: str = None, canonical_kind: str = None,
           sheet: str = None, only_uningested: bool = False) -> list:
    """Tìm trong catalog (không mở file). Lọc theo tên/công ty/canonical_kind/sheet."""
    q, cmp_ = _norm(query), _norm(company)
    ck, sh = _norm(canonical_kind), _norm(sheet)
    out = []
    for e in _load().values():
        if only_uningested and e.get("ingested"):
            continue
        if cmp_ and cmp_ not in _norm(e.get("company")):
            continue
        hay = _norm(e.get("file")) + " " + _norm(e.get("company")) + " " + _norm(e.get("report_type"))
        sheets_norm = " ".join(_norm(s["name"]) for s in e.get("sheets", []))
        cks = " ".join(_norm(s.get("canonical_kind")) for s in e.get("sheets", []) if s.get("canonical_kind"))
        if q and q not in hay and q not in sheets_norm:
            continue
        if ck and ck not in cks:
            continue
        if sh and sh not in sheets_norm:
            continue
        out.append(e)
    return out

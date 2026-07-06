# -*- coding: utf-8 -*-
"""MCP server luong 1 (ingest): template_analyze, discovery_record, import_plan_validate,
import_execute. Chay: `python -m servers.ingest_server` (stdio MCP server).

Tai su dung nguyen logic nghiep vu cua DashBoard_AI qua be_bridge - khong viet lai
detect/auto_map/commit/import_workbook."""
import hashlib
import io
import json
import os
import re

from mcp.server.fastmcp import FastMCP
from openpyxl import load_workbook

from .common import be_bridge as bb
from .common import canonical
from .common import contract
from .common import memory
from . import template_filler

mcp = FastMCP("dashboard_ingest")

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
_WORKSPACE_ROOT = os.path.dirname(_AGENT_ROOT)  # thư mục cha (vd d:\Company\ThinhCuong) - chứa
                                                  # DashBoard_Agent/DashBoard_AI/Data_test_dashboard
_IMPORTS_LEDGER = os.path.join(_AGENT_ROOT, "memory", "imports_ledger.json")


def _input_dir() -> str:
    input_dir = os.environ.get("INPUT_DIR") or "."
    return os.path.normpath(os.path.join(_AGENT_ROOT, input_dir))


def _read_file(file_path: str) -> bytes:
    if not os.path.isabs(file_path):
        file_path = os.path.join(_input_dir(), file_path)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")
    with open(file_path, "rb") as fh:
        return fh.read()


def _jsonable(v):
    """Excel cell value -> kiểu JSON-serializable an toàn (datetime/Decimal -> str)."""
    if v is None or isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def _all_sheets_columns(data: bytes) -> dict:
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        out = {}
        for ws in wb.worksheets:
            first_row = next(ws.iter_rows(max_row=1, values_only=True), [])
            out[ws.title] = [("" if c is None else str(c)) for c in first_row]
        return out
    finally:
        wb.close()


@mcp.tool()
def discover_files(root_dir: str = None, pattern: str = r".*\.xlsx$") -> dict:
    """Quét thư mục tìm workbook Excel — dùng khi chưa biết cần đọc file nào (nhiều
    công ty/nhiều file). RẺ: chỉ liệt kê tên file + đoán company/canonical_kind từ
    tên thư mục/tên file, KHÔNG mở nội dung workbook (mở file 17MB tốn 10-15s/lần,
    xem sheet_profile/generic_import_execute cho bước đọc nội dung).

    root_dir=None: mặc định quét INPUT_DIR. Có thể trỏ nơi khác nhưng PHẢI nằm trong
    workspace (cùng thư mục cha với DashBoard_Agent) — chặn quét ra ngoài phạm vi dự án.

    company_guess: nếu file nằm trong thư mục con của root_dir, dùng TÊN thư mục con đó;
    nếu không, thử regex trên tên file dạng 'X.N.MÃCTY.rest...' (vd 'B.7.AAG.TCKT...' -> 'AAG').
    canonical_kind_guess: đoán loại báo cáo từ TÊN FILE (xem servers/common/canonical.py) —
    file lớn/nhiều sheet (vd báo cáo tài chính riêng tổng hợp) thường không đoán được từ tên
    file mà phải gọi sheet_profile(sheet=None) để xem danh sách sheet bên trong."""
    base = os.path.normpath(os.path.join(_AGENT_ROOT, root_dir)) if root_dir else _input_dir()
    if os.path.commonpath([_WORKSPACE_ROOT, base]) != _WORKSPACE_ROOT:
        raise ValueError(f"'{root_dir}' nằm ngoài workspace — không được phép quét.")
    if not os.path.isdir(base):
        raise FileNotFoundError(f"Thư mục không tồn tại: {base}")

    rx = re.compile(pattern, re.IGNORECASE)
    code_rx = re.compile(r"^[A-Za-z]\.\d+\.([A-Za-z0-9]{2,8})\.")
    files = []
    for dirpath, _dirnames, filenames in os.walk(base):
        rel_dir = os.path.relpath(dirpath, base)
        for fn in filenames:
            if not rx.match(fn):
                continue
            if rel_dir == ".":
                company_guess = None
                m = code_rx.match(fn)
                if m:
                    company_guess = m.group(1)
            else:
                company_guess = rel_dir.split(os.sep)[0]
            rel_path = os.path.normpath(os.path.join(rel_dir, fn)) if rel_dir != "." else fn
            full_path = os.path.join(dirpath, fn)
            files.append({
                "path": rel_path.replace("\\", "/"),
                "file_name": fn,
                "company_guess": company_guess,
                "canonical_kind_guess": canonical.guess_canonical_kind(fn),
                "size_bytes": os.path.getsize(full_path),
            })
    companies = sorted({f["company_guess"] for f in files if f["company_guess"]})
    return {"root_dir": base, "companies": companies, "files": files}


@mcp.tool()
def sheet_profile(file_path: str, sheet: str = None, max_rows: int = 10,
                   max_cols: int = 8, col_depth: int = 30) -> dict:
    """Khám phá cấu trúc THÔ của 1 workbook/sheet khi template_analyze không nhận diện
    được report_type cố định nào (report_type=None). Dùng khi cần tự suy luận layout
    của 1 sheet lạ (khác 9 báo cáo THUCHI/SDT/HQKD/... đã biết).

    sheet=None: chỉ liệt kê tên các sheet trong workbook (rẻ, không đọc nội dung), kèm
    canonical_kind_guess đoán từ TÊN sheet (xem servers/common/canonical.py) để analyst
    biết ngay sheet nào đáng đào sâu trước khi đọc nội dung.
    sheet=<tên>: đọc 1 lượt sheet đó, trả 2 góc nhìn để tự nhận biết layout:
      - row_sample: max_rows dòng đầu, MỌI cột - phát hiện header nằm ngang/nhiều dòng
        (vd 2 dòng header như sheet '131'/'331': 'Đầu kỳ|Phát sinh|Cuối kỳ' rồi 'Nợ|Có').
      - col_sample: max_cols cột đầu, mỗi cột lấy col_depth dòng - phát hiện layout
        "thực thể nằm theo cột" (vd nhiều công ty làm cột như CĐKT hợp nhất).
      - merged_ranges: vùng ô merge (gợi ý ranh giới header nhiều dòng/nhiều cột).
    KHÔNG ghi gì, không cần DB - chỉ đọc file trong INPUT_DIR (qua _read_file)."""
    data = _read_file(file_path)
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        if sheet is None:
            return {
                "file_name": os.path.basename(file_path),
                "sheets": [{"sheet": s, "canonical_kind_guess": canonical.guess_canonical_kind(s)}
                           for s in wb.sheetnames],
            }

        if sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet}' không tồn tại. Các sheet có: {wb.sheetnames}")
        ws = wb[sheet]

        try:
            dimensions = ws.dimensions
        except Exception:
            dimensions = None

        try:
            merged_ranges = [str(r) for r in ws.merged_cells.ranges]
        except Exception:
            merged_ranges = []  # read_only worksheet có thể không có merged_cells

        depth = max(max_rows, col_depth)
        rows = []
        for row in ws.iter_rows(max_row=depth, values_only=True):
            rows.append([_jsonable(c) for c in row])

        row_sample = rows[:max_rows]
        col_sample = {}
        for ci in range(min(max_cols, max((len(r) for r in rows), default=0))):
            from openpyxl.utils import get_column_letter
            col_sample[get_column_letter(ci + 1)] = [
                (r[ci] if ci < len(r) else None) for r in rows[:col_depth]
            ]

        return {
            "file_name": os.path.basename(file_path), "sheet": sheet, "dimensions": dimensions,
            "merged_ranges": merged_ranges[:50], "row_sample": row_sample, "col_sample": col_sample,
            "n_rows_scanned": len(rows),
        }
    finally:
        wb.close()


@mcp.tool()
def sheet_routes(file_path: str) -> dict:
    """Try-route TẤT ĐỊNH cả workbook (mở 1 LẦN): phân loại MỖI sheet -> đích + trạng thái.
    Khác sheet_profile(sheet=None): CÓ đọc tiêu đề trong sheet để route theo NỘI DUNG (bắt tên
    sheet khó như '156'='BÁO CÁO NHẬP XUẤT TỒN', 'KHTS'='BẢNG TÍNH KHẤU HAO') và KHÔNG bỏ sót
    im lặng — mỗi sheet rơi vào đúng 1 nhóm:
      - routed        : target_sheet + canonical_kind + route_via ('name'|'content').
      - unknown       : có dữ liệu nhưng chưa nhận diện loại -> caller đẩy analyst/người.
      - skip_metadata : Master Data / User / DS báo cáo / BC THU CHI (TỔNG)...
      - empty         : không đủ dòng dữ liệu."""
    data = _read_file(file_path)
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        out = []
        for name in wb.sheetnames:
            rows = [row for row in wb[name].iter_rows(max_row=25, values_only=True)]
            non_empty = sum(1 for r in rows if any(c not in (None, "") for c in r))
            # Tiêu đề = ô ĐẦU TIÊN không rỗng của mỗi dòng trong ~8 dòng đầu (cột tiêu đề/merge),
            # KHÔNG lấy cả lưới -> tránh keyword vãng lai trong header/ô số khiến route nhầm.
            title_cells = []
            for r in rows[:8]:
                cell = next((c for c in r if c not in (None, "")), None)
                if cell is not None:
                    title_cells.append(str(cell))
            target, ck, via = contract.route_sheet(name, " ".join(title_cells))
            if via == "skip":
                status = "skip_metadata"
            elif target:
                status = "routed"
            elif non_empty >= 4:
                status = "unknown"
            else:
                status = "empty"
            out.append({"sheet": name, "canonical_kind": ck, "target_sheet": target,
                        "route_via": (via if via in ("name", "content") else None),
                        "status": status})
        return {"file_name": os.path.basename(file_path), "routes": out}
    finally:
        wb.close()


@mcp.tool()
def template_analyze(file_path: str) -> dict:
    """Phân tích 1 file xlsx chưa biết template: dò report_type/header/mapping,
    liệt kê toàn bộ sheet+cột, ghi vào discovery memory. Trả về TemplateSpec (dict)."""
    data = _read_file(file_path)
    file_name = os.path.basename(file_path)
    sheets_columns = _all_sheets_columns(data)

    try:
        prep = bb.prepare(data, file_name)
    except ValueError as e:
        # Không nhận diện được report_type - vẫn ghi discovery để qa biết đã từng thấy file này.
        record = memory.discovery_record(
            file_name=file_name, fingerprint="", sheets=list(sheets_columns.keys()),
            columns_per_sheet=sheets_columns, detected_report_type=None,
            anomalies=[str(e)],
        )
        return {
            "file_name": file_name, "report_type": None, "header_row": 0, "headers": [],
            "mapping": {}, "confidence": 0.0, "missing_required": [], "low_confidence_fields": [],
            "anomalies": [str(e)], "sample_rows": [], "all_sheets": list(sheets_columns.keys()),
        }

    confidence = 1.0 if prep["from_profile"] else (0.0 if prep["missing_required"] else 0.7)
    if prep["low_confidence"]:
        confidence = min(confidence, 0.5)
    fp = prep["fingerprint"]

    memory.discovery_record(
        file_name=file_name, fingerprint=fp, sheets=list(sheets_columns.keys()),
        columns_per_sheet=sheets_columns, detected_report_type=prep["report_type"],
        header_row=prep["header_row"], mapping=prep["mapping"], confidence=confidence,
        anomalies=([f"Thiếu field bắt buộc: {prep['missing_required']}"] if prep["missing_required"] else []),
    )

    return {
        "file_name": file_name,
        "report_type": prep["report_type"],
        "header_row": prep["header_row"],
        "headers": prep["headers"],
        "mapping": prep["mapping"],
        "confidence": confidence,
        "missing_required": prep["missing_required"],
        "low_confidence_fields": prep["low_confidence"],
        "anomalies": [],
        "sample_rows": prep["sample_rows"],
        "all_sheets": list(sheets_columns.keys()),
        "fingerprint": fp,
    }


@mcp.tool()
def discovery_record_tool(
    file_name: str, fingerprint: str, sheets: list, columns_per_sheet: dict,
    detected_report_type: str = None, header_row: int = None, mapping: dict = None,
    period: str = None, confidence: float = 0.0, anomalies: list = None,
) -> dict:
    """Ghi thủ công 1 bản ghi vào discovery memory (dùng khi agent muốn bổ sung
    ghi chú/period sau khi đã template_analyze)."""
    return memory.discovery_record(
        file_name, fingerprint, sheets, columns_per_sheet, detected_report_type,
        header_row, mapping, period, confidence, anomalies,
    )


@mcp.tool()
def import_plan_validate(
    file_name: str, report_type: str, mapping: dict, period: str = None,
    dataset_kind: str = "day",
) -> dict:
    """Kiểm ImportPlan trước khi ghi: field bắt buộc theo FIELD_DEFS, report_type hợp lệ,
    period hợp lệ (YYYY-MM) nếu kind=month. KHÔNG ghi DB."""
    errors, warnings = [], []
    if report_type not in bb.FIELD_DEFS:
        errors.append(f"report_type '{report_type}' không hợp lệ (phải là 1 trong {bb.REPORT_CODES}).")
        return {"ok": False, "errors": errors, "warnings": warnings}

    fields = bb.FIELD_DEFS[report_type]
    missing = [f[0] for f in fields if f[2] and f[0] not in mapping]
    if missing:
        errors.append(f"Thiếu mapping cho field bắt buộc: {missing}")

    if dataset_kind == "month":
        import re
        if not period or not re.match(r"^\d{4}-\d{2}$", period):
            warnings.append("period không đúng định dạng YYYY-MM hoặc bị thiếu (sẽ tự suy từ tên file/dữ liệu).")
    elif dataset_kind not in ("day", "month"):
        errors.append("dataset_kind phải là 'day' hoặc 'month'.")

    return {"ok": not errors, "errors": errors, "warnings": warnings}


def _ledger_load() -> list:
    if not os.path.exists(_IMPORTS_LEDGER):
        return []
    try:
        with open(_IMPORTS_LEDGER, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return []


def _ledger_save(entries: list):
    os.makedirs(os.path.dirname(_IMPORTS_LEDGER), exist_ok=True)
    with open(_IMPORTS_LEDGER, "w", encoding="utf-8") as fh:
        json.dump(entries, fh, ensure_ascii=False, indent=2)


_ENTITY_CODE_HINTS = ("ma_doi_tuong", "ma_dt", "ma_kh", "ma_ncc", "ma_khach", "code", "ma")
_ENTITY_NAME_HINTS = ("ten_doi_tuong", "ten_dt", "ten_kh", "ten_ncc", "ten_khach", "name", "ten", "dien_giai")


def _normalize_columns(columns: list) -> list:
    """Chuẩn hoá 'columns' về schema ColumnRole {index, role, label} — khoan dung với biến thể
    mà LLM hay sinh: 'col_letter' (A/B/..) -> index; 'field_name' -> role nếu thiếu 'role'.
    Nếu role không phải từ khoá đặc biệt nhưng field_name gợi ý mã/tên đối tượng thì map về
    entity_code/entity_name (nếu chưa có cột nào giữ vai trò đó) để melt gắn đúng thực thể."""
    from openpyxl.utils import column_index_from_string

    norm = []
    have_code = have_name = False
    for c in columns:
        idx = c.get("index")
        if idx is None and c.get("col_letter"):
            idx = column_index_from_string(str(c["col_letter"]).strip().upper()) - 1
        if idx is None:
            continue
        role = c.get("role") or c.get("field_name") or ""
        label = c.get("label", "")
        norm.append({"index": int(idx), "role": role, "label": label})

    reserved = {"entity_code", "entity_name", "label", "skip"}
    have_code = any(c["role"] == "entity_code" for c in norm)
    have_name = any(c["role"] == "entity_name" for c in norm)
    for c in norm:
        if c["role"] in reserved:
            continue
        low = c["role"].lower()
        if not have_code and any(low == h or low.startswith(h) for h in _ENTITY_CODE_HINTS):
            c["role"] = "entity_code"; have_code = True
        elif not have_name and any(low == h or low.startswith(h) for h in _ENTITY_NAME_HINTS):
            c["role"] = "entity_name"; have_name = True
    return norm


def _ledger_find(dataset_kind, report_type, period, fingerprint, content_hash=None):
    """Tìm bản ghi import trùng. Khoá dedup = (dataset_kind, content_hash) — content_hash
    (sha1 nội dung file) định danh file duy nhất: CÙNG nội dung file -> coi là trùng, bỏ
    qua (idempotent thật sự); KHÁC nội dung dù cùng layout/report_type/period vẫn được nạp
    (fix S1). Không so khớp period/fingerprint để tránh lệch None-vs-suy-diễn ở kind=month;
    2 tham số đó vẫn được lưu làm metadata."""
    if not content_hash:
        return None
    for e in _ledger_load():
        if e.get("dataset_kind") == dataset_kind and e.get("content_hash") == content_hash:
            return e
    return None


@mcp.tool()
def import_execute(
    file_path: str, dataset_kind: str = "day", report_type: str = None,
    mapping: dict = None, period: str = None, dry_run: bool = True,
) -> dict:
    """Ghi dữ liệu tinh vào raw_rows (idempotent - kiểm memory/imports_ledger.json
    theo khoá (dataset_kind, report_type, period, fingerprint) trước khi ghi thật).

    dry_run=True (mặc định): chỉ preview số dòng dự kiến, KHÔNG ghi DB.
    dataset_kind='day' -> tái dùng importer.prepare/commit (9 báo cáo, cần report_type+mapping).
    dataset_kind='month' -> tái dùng importer_month.import_workbook (nhiều sheet, tự suy period).
    """
    data = _read_file(file_path)
    file_name = os.path.basename(file_path)

    if dataset_kind == "month":
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        try:
            is_month = bb.detect_ledger(wb)
        finally:
            wb.close()
        if not is_month:
            return {"dry_run": dry_run, "by_type": {}, "skipped_duplicate": False,
                    "message": f"'{file_name}' không phải template Tháng (thiếu sheet DATA dạng sổ cái)."}
        fp = bb.fingerprint([file_name])  # fingerprint theo tên file cho kind=month (không có header cột chung)
        content_hash = hashlib.sha1(data).hexdigest()
        dup = _ledger_find("month", "MONTH", period, fp, content_hash)
        if dup and not dry_run:
            return {"dry_run": False, "dataset_id": dup["dataset_id"], "by_type": {},
                    "skipped_duplicate": True,
                    "message": f"Đã import trước đó (dataset_id={dup['dataset_id']}, period={dup.get('period')}). Bỏ qua để tránh trùng."}
        if dry_run:
            return {"dry_run": True, "by_type": {}, "skipped_duplicate": bool(dup),
                    "message": "Dry-run: sẽ gọi importer_month.import_workbook (số dòng thực tế biết sau khi ghi)."}
        ds = bb.new_dataset(kind="month", period=period, name=file_name)
        result = bb.import_workbook(ds["id"], data, file_name)
        if result.get("period"):
            bb.repo.set_period(ds["id"], result["period"])
        entries = _ledger_load()
        entries.append({"dataset_kind": "month", "report_type": "MONTH", "period": result.get("period") or period,
                         "fingerprint": fp, "content_hash": content_hash, "dataset_id": ds["id"], "file_name": file_name})
        _ledger_save(entries)
        return {"dry_run": False, "dataset_id": ds["id"], "by_type": result["by_type"],
                "skipped_duplicate": False, "message": "Đã ghi raw_rows.", "period": result.get("period")}

    # dataset_kind == "day"
    try:
        prep = bb.prepare(data, file_name)
    except ValueError as e:
        return {"dry_run": dry_run, "by_type": {}, "skipped_duplicate": False, "message": str(e)}

    final_map = dict(prep["mapping"])
    if mapping:
        final_map.update({k: int(v) for k, v in mapping.items() if v is not None})
    rt = report_type or prep["report_type"]
    fp = prep["fingerprint"]
    content_hash = hashlib.sha1(data).hexdigest()

    dup = _ledger_find("day", rt, period, fp, content_hash)
    if dup and not dry_run:
        return {"dry_run": False, "dataset_id": dup["dataset_id"], "by_type": {},
                "skipped_duplicate": True,
                "message": f"File cùng report_type+fingerprint đã import trước đó (dataset_id={dup['dataset_id']}). Bỏ qua để tránh trùng."}

    if dry_run:
        n_rows = sum(1 for r in prep["rows"][prep["header_row"] + 1:] if r and any(c is not None for c in r))
        return {"dry_run": True, "by_type": {rt: n_rows}, "skipped_duplicate": bool(dup),
                "message": f"Dry-run: dự kiến ghi ~{n_rows} dòng report_type={rt}."}

    ds = bb.new_dataset(kind="day", period=period, name=file_name)
    result = bb.commit(ds["id"], rt, prep["header_row"], prep["headers"], final_map, prep["rows"])
    entries = _ledger_load()
    entries.append({"dataset_kind": "day", "report_type": rt, "period": period,
                     "fingerprint": fp, "content_hash": content_hash, "dataset_id": ds["id"], "file_name": file_name})
    _ledger_save(entries)
    return {"dry_run": False, "dataset_id": ds["id"], "by_type": {rt: result["rows"]},
            "skipped_duplicate": False, "message": "Đã ghi raw_rows.", "issues": result["issues"]}


@mcp.tool()
def generic_import_execute(
    file_path: str, sheet: str, mapping: dict, period: str = None, ngay: str = None,
    cong_ty: str = None, dry_run: bool = True,
) -> dict:
    """Ghi dữ liệu theo 1 SheetMapping khai báo (xem servers/common/models.py SheetMapping) —
    dùng cho sheet KHÔNG khớp 9 report_type cố định (vd '131'/'331'/'Biểu khấu hao' trong file
    báo cáo tài chính riêng). KHÔNG sinh/chạy code tuỳ ý — chỉ diễn giải 'columns' (row_major)
    hoặc 'entities'+'row_roles' (column_major) theo đúng 1 thuật toán cố định: "melt" bảng rộng
    thành raw_rows dạng dài (tidy), tổng quát hoá pattern importer_month._parse_cdkt.

    mapping (dict) tối thiểu cần: orientation ('row_major'|'column_major'),
    target_report_type (PHẢI có tiền tố 'GEN_'), data_start_row, và columns[] hoặc
    entities[]+row_roles[] tương ứng orientation (xem models.SheetMapping). Có thể kèm
    canonical_kind (vd 'TK131') để report_spec_search tìm lại được ở công ty khác.

    cong_ty: tên công ty/đơn vị sở hữu file này (từ discover_files hoặc SheetMapping.company)
    — CHỈ dùng cho orientation='row_major' (mỗi file thường thuộc 1 công ty); với
    'column_major' công ty đã lấy từ tên thực thể (entities[].entity_name) nên bỏ qua tham số này.

    LUÔN trả sample_mapped_rows (5 dòng đầu đã map) bất kể dry_run=True/False, để người dùng
    soi trước khi orchestrator cho approve ghi thật (suy luận LLM, rủi ro map sai cột cao hơn
    auto_map cũ)."""
    target_rt = mapping.get("target_report_type") or ""
    if not target_rt.startswith("GEN_"):
        raise ValueError("target_report_type phải có tiền tố 'GEN_' để tách biệt report_type cố định.")
    orientation = mapping.get("orientation")
    if orientation not in ("row_major", "column_major"):
        raise ValueError("orientation phải là 'row_major' hoặc 'column_major'.")

    data = _read_file(file_path)
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet}' không tồn tại. Các sheet có: {wb.sheetnames}")
        rows = [[_jsonable(c) for c in r] for r in wb[sheet].iter_rows(values_only=True)]
    finally:
        wb.close()

    file_name = os.path.basename(file_path)
    data_start = mapping.get("data_start_row", 0)
    ngay_val = ngay or (f"{period}-01" if period and len(period) == 7 else None)
    period_val = period or (ngay_val[:7] if ngay_val else None)

    # (row_index_nguồn, entity_code, entity_name, role, khoi, cong_ty, amount, payload)
    melted = []
    if orientation == "row_major":
        cols = {c["index"]: c for c in _normalize_columns(mapping.get("columns", []))}
        code_idx = next((i for i, c in cols.items() if c["role"] == "entity_code"), None)
        name_idx = next((i for i, c in cols.items() if c["role"] == "entity_name"), None)
        for ri in range(data_start, len(rows)):
            row = rows[ri]
            if not row or all(v is None for v in row):
                continue
            entity_code = bb.parse_text(row[code_idx]) if code_idx is not None and code_idx < len(row) else None
            entity_name = bb.parse_text(row[name_idx]) if name_idx is not None and name_idx < len(row) else None
            row_payload = {c["role"]: bb.parse_text(row[idx]) for idx, c in cols.items()
                           if idx < len(row) and row[idx] is not None}
            for idx, c in cols.items():
                role = c["role"]
                if role in ("entity_code", "entity_name", "label", "skip") or idx >= len(row):
                    continue
                val = bb.parse_num(row[idx])
                if not val:
                    continue
                melted.append((ri, entity_code, entity_name, role, None, cong_ty, val, row_payload))
    else:  # column_major
        for rr in mapping.get("row_roles", []):
            ri = rr["row_index"]
            if ri >= len(rows):
                continue
            row = rows[ri]
            for ent in mapping.get("entities", []):
                ci = ent["col_index"]
                if ci >= len(row) or row[ci] is None:
                    continue
                val = bb.parse_num(row[ci])
                if not val:
                    continue
                payload = {"row_label": rr.get("label", ""), "entity": ent["entity_name"]}
                melted.append((ri, rr["role"], rr.get("label", ""), None,
                               ent["entity_name"], ent["entity_name"], val, payload))

    sample_mapped_rows = [
        {"source_row": m[0], "dim1": m[1], "dim2": m[2], "dim3": m[3],
         "khoi": m[4], "cong_ty": m[5], "amount": m[6]}
        for m in melted[:5]
    ]

    fp = bb.fingerprint([file_name, sheet, orientation,
                         json.dumps(mapping.get("columns") or mapping.get("row_roles") or [],
                                    ensure_ascii=False)])

    if dry_run:
        # Khép vòng học sớm: dry-run thành công (có dòng melt được) -> lưu mapping ngay
        # (verified=False), không cần chờ ai bấm "ghi thật". Nhờ vậy autobatch/propose không
        # phải chạy lại LLM cho cùng 1 sheet đã verify đúng qua dry-run trước đó.
        if melted:
            memory.report_spec_save(fp, mapping, verified=False)
        return {"dry_run": True, "target_report_type": target_rt, "row_count": len(melted),
                "sample_mapped_rows": sample_mapped_rows, "skipped_duplicate": False,
                "message": f"Dry-run: dự kiến ghi {len(melted)} dòng report_type={target_rt}."}

    content_hash = hashlib.sha1(data).hexdigest()
    dup = _ledger_find("generic", target_rt, period_val, fp, content_hash)
    if dup:
        return {"dry_run": False, "dataset_id": dup["dataset_id"], "row_count": 0,
                "sample_mapped_rows": sample_mapped_rows, "skipped_duplicate": True,
                "message": f"Sheet này (cùng mapping) đã nạp trước đó, dataset_id={dup['dataset_id']}."}

    ds = bb.new_dataset(kind="generic", period=period_val, name=f"{file_name}::{sheet}")
    to_insert = [
        (ds["id"], target_rt, ri, ngay_val, cong_ty, khoi, None, period_val, val, None,
         entity_code, entity_name, role, json.dumps(payload, ensure_ascii=False))
        for ri, entity_code, entity_name, role, khoi, cong_ty, val, payload in melted
    ]
    db = bb.get_db()
    db.executemany(
        "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,"
        "cost_center,period_month,amount,amount2,dim1,dim2,dim3,payload) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        to_insert,
    )
    db.commit()

    entries = _ledger_load()
    entries.append({"dataset_kind": "generic", "report_type": target_rt, "period": period_val,
                     "fingerprint": fp, "content_hash": content_hash, "dataset_id": ds["id"],
                     "file_name": file_name, "sheet": sheet})
    _ledger_save(entries)
    memory.report_spec_save(fp, mapping, verified=True)

    return {"dry_run": False, "dataset_id": ds["id"], "row_count": len(to_insert),
            "sample_mapped_rows": sample_mapped_rows, "skipped_duplicate": False,
            "message": f"Đã ghi {len(to_insert)} dòng report_type={target_rt}."}


@mcp.tool()
def template_contract_info() -> dict:
    """MÔ TẢ TEMPLATE VÀNG cho analyst — ĐỌC TRƯỚC KHI DỰNG MAPPING.

    Trả bản guide đầy đủ: đơn vị (tỷ đồng -> value_scale khi nguồn VND), quy tắc map,
    mỗi sheet {mục đích, grain, cot_nhap_lieu (CHỈ map vào đây), cot_auto_bo_qua (công thức, bỏ)},
    danh sách công ty/khối/mã cost center hợp lệ (để đặt constants), tên chỉ tiêu KQKD chuẩn,
    và sheet đích theo loại báo cáo nguồn.
    """
    return contract.guide()


@mcp.tool()
def template_fill(file_path: str, source_sheet: str, target_sheet: str, mapping: dict,
                  period: str = None, cong_ty: str = None, dry_run: bool = True,
                  auto_import: bool = False, value_scale: float = 1.0,
                  constants: dict = None, normalize_kqkd: bool = True,
                  rename_rows: dict = None) -> dict:
    """Điền số liệu từ file nguồn vào BẢN SAO template vàng, theo mapping cột.

    - mapping: {tên cột TEMPLATE (input): tên cột NGUỒN}. Xem template_contract_info() để biết cột.
    - constants: {tên cột TEMPLATE: giá trị HẰNG} — gán cứng mọi dòng (vd cost center khi sheet nguồn
      cấp CÔNG TY không có CC theo dòng: constants={'Mã Cost center ◀ NHẬP':'ST_GD'}).
    - rename_rows: {tên chỉ tiêu NGUỒN: tên chỉ tiêu CHUẨN} — ĐỔI TÊN đúng vài dòng về tên chuẩn để KPI
      sáng (KHÁC mapping: mapping đổi CỘT; rename_rows đổi GIÁ TRỊ tên của DÒNG). BẮT BUỘC cho 01_HQKD khi
      nguồn đặt tên riêng: chọn ĐÚNG 1 dòng tổng doanh thu -> 'Doanh thu thuần', 1 dòng tổng chi phí ->
      'Tổng chi phí', 1 dòng lợi nhuận trước thuế -> 'Lợi nhuận trước thuế'. Khớp KHÔNG DẤU + chứa-trong;
      CHỌN 1 DÒNG TỔNG CẤP CAO NHẤT / chỉ tiêu để KHÔNG cộng trùng với dòng con.
    - value_scale: nhân cột tiền '(tỷ)'. NGUỒN VND -> template TỶ ĐỒNG phải dùng value_scale=1e-9.
      (tool KHÔNG tự đổi đơn vị — phải truyền value_scale đúng, nếu không số sai 1 tỷ lần.)
    - Cost center resolve về mã chuẩn; không khớp -> unmapped_cc + để nguyên (khối=NULL -> "(Chưa phân bổ)").
    - dry_run=True: preview. dry_run=False: ghi file điền -> out_path; auto_import=True: nạp raw_rows.
    """
    data = _read_file(file_path)
    return template_filler.fill_from_source(
        data, source_sheet, target_sheet, mapping, period=period, cong_ty=cong_ty,
        file_name=os.path.basename(file_path), dry_run=dry_run, auto_import=auto_import,
        value_scale=value_scale, constants=constants, normalize_kqkd=normalize_kqkd,
        rename_rows=rename_rows, source_path=file_path)


@mcp.tool()
def catalog_reindex(root_dir: str = None) -> dict:
    """Quét lại thư mục file đã kéo về (Connect_VPS/received_reports mặc định) -> cập nhật
    source_catalog. Trả {indexed, total_in_catalog}. QA tra qua catalog_search."""
    from .common import source_catalog
    return source_catalog.index_dir(root_dir)


@mcp.tool()
def template_import(filled_path: str, cong_ty: str = None) -> dict:
    """Nạp 1 file ĐÃ ĐIỀN template chuẩn (do template_fill tạo trong template_trust/filled/) vào
    raw_rows. cong_ty: GỘP đa-công-ty theo kỳ (1 dataset/kỳ, idempotent theo (kỳ,cong_ty), đóng dấu
    cong_ty). Trả {dataset_id, grain, by_type, rows_imported, period}."""
    if not os.path.isabs(filled_path):
        filled_path = os.path.join(template_filler.FILLED_DIR, filled_path)
    return template_filler.import_filled(filled_path, cong_ty=cong_ty)


if __name__ == "__main__":
    mcp.run()

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
import re
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
    """Quét thư mục received_reports, index mọi .xlsx (bỏ file tạm ~$). Đuôi so KHÔNG phân biệt
    hoa/thường — SRVF gửi '.Xlsx' (2026-07-30), glob '*.xlsx' bỏ sót -> file nằm trên đĩa nhưng
    catalog không thấy: tab Nguồn dữ liệu báo 'Mới · chưa kéo về', không analyze/xem được."""
    root = root or RECEIVED_DIR
    if not os.path.isdir(root):
        return {"ok": False, "error": f"Không thấy thư mục: {root}", "indexed": 0}
    files = [f for f in glob.glob(os.path.join(root, "**", "*"), recursive=True)
             if f.lower().endswith(".xlsx") and not os.path.basename(f).startswith("~$")]
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


_DAU_PHAN_CACH = re.compile(r"[^a-z0-9]+")


def _dac(s) -> str:
    """Bỏ SẠCH mọi ký tự không phải chữ/số: 'B.9.TC.TCKT.M.202607' -> 'b9tctcktm202607'.

    Tên file nguồn KHÔNG nhất quán dấu chấm — 10/11 bản báo cáo tài chính riêng tháng 7 ghi
    'M.202607' còn XDV ghi 'M202607'. So chuỗi con nguyên văn thì hỏi 'M202607' chỉ ra ĐÚNG 1
    file XDV, 10 file kia tàng hình, và agent QA đã kết luận "toàn tập đoàn chỉ có 1 báo cáo"
    rồi lấy luôn lợi nhuận riêng XDV làm lợi nhuận tập đoàn (12/08/2026). Ép cả hai vế về dạng
    đặc rồi mới so thì dấu phân cách không còn quyết định tìm thấy hay không.
    """
    return _DAU_PHAN_CACH.sub("", _norm(s))


# Nguồn TỔNG HỢP nhưng chỉ ở phạm vi MỘT khối/pháp nhân. Model đọc file thấy chữ "hợp nhất" là
# tưởng số cấp tập đoàn — đã trả nhầm lợi nhuận khối Xe tải (15,05 tỷ / 2,43 tỷ) thành lợi nhuận
# toàn tập đoàn ba lần liên tiếp, kể cả sau khi SKILL nêu đích danh cái bẫy này. Nhắc trong prompt
# không ăn thua vì prompt dài và cảnh báo ở xa dữ liệu; gắn thẳng vào KẾT QUẢ TOOL thì model đọc
# cùng lúc với con số. Khớp theo report_type nên file mới cùng loại tự có cảnh báo.
_PHAM_VI_HEP = {
    "baocaotaichinhhopnhatxetai":
        "CHỈ hợp nhất KHỐI XE TẢI của pháp nhân HT — KHÔNG phải số toàn tập đoàn. "
        "Sheet 'kqkd tổng hợp nhất' / 'BCĐKT hợp nhất' ở đây đều là phạm vi khối xe tải.",
}
# Sheet mang chữ "hợp nhất" nhưng nằm trong báo cáo RIÊNG của một pháp nhân (vd 'HQKD HỢP NHẤT GA').
_SHEET_HEP = "hop nhat"


def _canh_bao_pham_vi(e: dict):
    """Câu cảnh báo phạm vi cho một mục catalog, hoặc None nếu không có gì phải cảnh báo."""
    rt = _norm(e.get("report_type"))
    if rt in _PHAM_VI_HEP:
        return _PHAM_VI_HEP[rt]
    ten = [s["name"] for s in e.get("sheets", []) if _SHEET_HEP in _norm(s["name"])]
    if ten:
        return (f"Sheet {ten} có chữ 'hợp nhất' nhưng đây là báo cáo RIÊNG của "
                f"{e.get('company') or 'một pháp nhân'} — phạm vi một pháp nhân, KHÔNG phải tập đoàn.")
    return None


def search(query: str = None, company: str = None, canonical_kind: str = None,
           sheet: str = None, only_uningested: bool = False,
           month=None, report_type: str = None) -> list:
    """Tìm trong catalog (không mở file). Lọc theo tên/công ty/kỳ/loại báo cáo/sheet.

    `query` tách thành TỪ KHOÁ theo khoảng trắng và phải khớp HẾT (AND), mỗi từ khoá khớp khi
    xuất hiện trong tên-file-đã-bỏ-dấu-phân-cách (xem `_dac`). Trước đây là một phép `in` nguyên
    văn: câu hỏi 2 từ trở lên gần như luôn rỗng, mà 1 từ thì lệch một dấu chấm là trượt.

    `month` / `report_type` là đường TẤT ĐỊNH, nên dùng thay vì nhét kỳ vào `query`: kỳ nằm
    trong tên file dưới nhiều dạng ('M.202607', 'M202607', 'M.2026.07'), dò bằng chuỗi là may rủi.
    """
    cmp_ = _norm(company)
    ck, sh = _norm(canonical_kind), _norm(sheet)
    rt = _norm(report_type)
    tu_khoa = [_dac(t) for t in (query or "").split() if _dac(t)]
    try:
        thang = int(month) if month not in (None, "") else None
    except (TypeError, ValueError):
        thang = None
    out = []
    for e in _load().values():
        if only_uningested and e.get("ingested"):
            continue
        if cmp_ and cmp_ not in _norm(e.get("company")):
            continue
        if thang is not None and e.get("month") != thang:
            continue
        if rt and rt not in _norm(e.get("report_type")):
            continue
        hay = _dac(e.get("file")) + " " + _dac(e.get("company")) + " " + _dac(e.get("report_type"))
        sheets_norm = " ".join(_norm(s["name"]) for s in e.get("sheets", []))
        sheets_dac = " ".join(_dac(s["name"]) for s in e.get("sheets", []))
        cks = " ".join(_norm(s.get("canonical_kind")) for s in e.get("sheets", []) if s.get("canonical_kind"))
        if tu_khoa and not all(t in hay or t in sheets_dac for t in tu_khoa):
            continue
        if ck and ck not in cks:
            continue
        if sh and sh not in sheets_norm:
            continue
        cb = _canh_bao_pham_vi(e)
        out.append({**e, "canh_bao_pham_vi": cb} if cb else e)
    return out

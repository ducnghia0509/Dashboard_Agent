# -*- coding: utf-8 -*-
"""MCP server luong 2 (qa): sql_query, glossary_lookup, discovery_search, source_inspect.
Chay: `python -m servers.qa_server` (stdio MCP server).

KHONG dung RAG/vector - moi tra cuu la deterministic (SQL that, JSON tra cuu chinh xac)."""
import json
import os

from mcp.server.fastmcp import FastMCP

from .common import be_bridge as bb
from .common import guardrails, introspect, memory
from .common.db_ro import get_ro_db

mcp = FastMCP("dashboard_qa")

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
_KPI_GLOSSARY_PATH = os.path.join(_AGENT_ROOT, "kpi_glossary.json")
# Tài liệu tham chiếu (mapping/org chart) - CHỈ ĐỌC, không phải file nguồn báo cáo. Thêm 2026-07-30
# để qa đọc thẳng Tài liệu/Mapping_Dashboard_QTTC.xlsx (cách tính 50 chỉ tiêu theo từng công ty)
# thay vì phải hardcode lại nội dung vào SKILL.md.
_DOCS_DIR = os.path.normpath(os.path.join(_AGENT_ROOT, "..", "Tài liệu"))

_kpi_cache = None


def _kpi_glossary() -> list:
    global _kpi_cache
    if _kpi_cache is None:
        if os.path.exists(_KPI_GLOSSARY_PATH):
            with open(_KPI_GLOSSARY_PATH, encoding="utf-8") as fh:
                _kpi_cache = json.load(fh)
        else:
            _kpi_cache = []
    return _kpi_cache


def _norm(s) -> str:
    return bb.normalize_header(s or "", True)


# TẠM KHÓA (2026-07-24): chặn hẳn đường SQL/Postgres của luồng qa — bắt buộc qa agent
# dùng source_inspect/catalog_search đọc thẳng file Excel trong Connect_VPS/received_reports
# thay vì truy vấn raw_rows. Bỏ comment 2 tool dưới (+ restart dashboard-mcp-qa.service) nếu
# thực sự cần dùng lại sql_query.
#
# @mcp.tool()
# def sql_query(sql: str, params: list = None) -> dict:
#     """Chạy 1 câu SELECT/WITH read-only trên raw_rows (Postgres, role read-only).
#     Guardrails tự chặn DML/DDL/multi-statement và ép LIMIT. Trả rows + sql_executed."""
#     conn = get_ro_db()
#     rows, safe_sql = guardrails.run_readonly(conn, sql, params)
#     return {"rows": rows, "sql_executed": safe_sql, "row_count": len(rows)}
#
#
# @mcp.tool()
# def schema_describe() -> str:
#     """Mô tả cấu trúc bảng raw_rows + report_type hợp lệ (dùng để tự sinh SQL)."""
#     return introspect.schema_describe()


@mcp.tool()
def glossary_lookup(term: str) -> dict:
    """Tra cứu định nghĩa/công thức/nguồn dữ liệu theo từ khoá (không dấu, không phân
    biệt hoa/thường) trong: master_data (công ty/khối/cost center/chiều phân tích),
    FIELD_DEFS/FIELD_LABELS/REPORT_LABELS (schema 9 báo cáo), và kpi_glossary.json
    (50 chỉ số quản trị từ guideline.xlsx, kèm công thức + cảnh báo đỏ + nguồn)."""
    nt = _norm(term)
    out = {"kpi_glossary": [], "field_defs": [], "master_data": [], "report_types": []}

    for rec in _kpi_glossary():
        hay = _norm(" ".join([
            rec.get("chi_tieu", ""), rec.get("nhom_bao_cao", ""), rec.get("nhom_con", ""),
            rec.get("chieu_phan_tich", ""), rec.get("canh_bao_do", ""),
            rec.get("nguon_du_lieu", ""), rec.get("cong_thuc", ""),
            # ghi_chu/y_nghia chứa cảnh báo "CHƯA CÓ NGUỒN" — phải index để QA
            # tìm ra cảnh báo nguồn thay vì khẳng định có số (xem gen_kpi_glossary).
            rec.get("ghi_chu", ""), rec.get("y_nghia", ""), rec.get("canh_bao_nguon", ""),
        ]))
        if nt in hay:
            out["kpi_glossary"].append(rec)

    for rt, fields in bb.FIELD_DEFS.items():
        for key, typ, required, _idx, _extra in fields:
            label = bb.FIELD_LABELS.get(key, key)
            if nt in _norm(key) or nt in _norm(label):
                out["field_defs"].append({
                    "report_type": rt, "report_label": bb.REPORT_LABELS.get(rt, rt),
                    "key": key, "label": label, "type": typ, "required": required,
                })

    for code, label in bb.REPORT_LABELS.items():
        if nt in _norm(code) or nt in _norm(label):
            out["report_types"].append({"code": code, "label": label})

    md = bb.master_data()
    for section in ("companies", "khoi", "costCenters", "chieuPhanTich"):
        for item in md.get(section, []):
            text = json.dumps(item, ensure_ascii=False)
            if nt in _norm(text):
                out["master_data"].append({"section": section, "item": item})

    out["total_matches"] = sum(len(v) for v in out.values() if isinstance(v, list))
    return out


@mcp.tool()
def discovery_search(query: str = None, report_type: str = None) -> dict:
    """Tìm trong discovery memory: số này/file này từng được phân tích chưa, đến từ
    sheet/cột nào, report_type gì, mapping ra sao. Trả {"results": [...]}."""
    return {"results": memory.discovery_search(query=query, report_type=report_type)}


@mcp.tool()
def report_spec_search(query: str = None, sheet: str = None, target_report_type: str = None,
                        canonical_kind: str = None) -> dict:
    """Tìm trong catalog SheetMapping đã học (Extension 2 - sheet lạ không khớp 9 report_type
    cố định, vd '131'/'331'/'Biểu khấu hao'): sheet này đã có cách lấy dữ liệu (mapping) chưa,
    report_type GEN_* này lấy từ sheet/cột nào. canonical_kind (vd 'TK131') tìm được mapping
    đã học ở CÔNG TY KHÁC dù tên sheet/file khác nhau, miễn cùng loại báo cáo. Dùng trước khi
    phân tích lại từ đầu bằng sheet_profile, và dùng làm ngữ cảnh khi qa cần giải thích 1 số
    liệu GEN_*. Trả {"results": [...]}."""
    return {"results": memory.report_spec_search(query=query, sheet=sheet,
                                                   target_report_type=target_report_type,
                                                   canonical_kind=canonical_kind)}


def _input_dir() -> str:
    input_dir = os.environ.get("INPUT_DIR") or "../Data_test_dashboard"
    return os.path.normpath(os.path.join(_AGENT_ROOT, input_dir))


def _within(base: str, target: str) -> bool:
    try:
        return os.path.commonpath([base, target]) == base
    except ValueError:
        return False


def _allowed_read_dirs() -> list:
    """Thư mục source_inspect được phép đọc: INPUT_DIR (template chuẩn) +
    Connect_VPS/received_reports (file gốc kéo về; catalog_search trả path tuyệt đối ở đây) +
    Tài liệu (tài liệu tham chiếu/mapping, KHÔNG phải file nguồn báo cáo - chỉ đọc)."""
    from .common.source_catalog import RECEIVED_DIR
    return [_input_dir(), RECEIVED_DIR, _DOCS_DIR]


def _resolve_readable(file_name: str) -> str:
    """Phân giải file_name (tên trơn / tương đối / tuyệt đối) về đường dẫn nằm TRONG một
    thư mục được phép. Chặn path traversal; raise nếu ngoài phạm vi hoặc không tồn tại.

    Model hay gọi CHỈ tên file trơn (không kèm thư mục con, vd company/report_type) thay vì
    path tương đối đầy đủ hoặc path tuyệt đối lấy từ catalog_search - fallback: tìm đệ quy theo
    basename trong các thư mục được phép, chỉ nhận nếu khớp DUY NHẤT (tránh trả nhầm file trùng
    tên ở công ty khác)."""
    bases = _allowed_read_dirs()
    if os.path.isabs(file_name):
        targets = [os.path.normpath(file_name)]
    else:
        targets = [os.path.normpath(os.path.join(b, file_name)) for b in bases]
    in_scope = [t for t in targets if any(_within(b, t) for b in bases)]
    if not in_scope:
        raise ValueError(f"'{file_name}' nằm ngoài thư mục được phép đọc "
                         f"(INPUT_DIR / Connect_VPS/received_reports / Tài liệu) - không được phép.")
    for t in in_scope:
        if os.path.exists(t):
            return t

    base_name = os.path.basename(file_name)
    matches = []
    for b in bases:
        if not os.path.isdir(b):
            continue
        for root, _dirs, files in os.walk(b):
            if base_name in files:
                matches.append(os.path.join(root, base_name))
    unique = sorted(set(matches))
    if len(unique) == 1:
        return unique[0]
    if len(unique) > 1:
        raise FileNotFoundError(
            f"'{file_name}' khớp NHIỀU file trùng tên trong INPUT_DIR/received_reports/Tài liệu "
            f"- dùng path tương đối đầy đủ (vd từ catalog_search) để chọn đúng: {unique}")
    raise FileNotFoundError(f"Không tìm thấy '{file_name}' trong INPUT_DIR, received_reports hoặc Tài liệu.")


# Nhóm báo cáo chứa dữ liệu cá nhân. `qa` phục vụ nhiều người dùng dashboard và định tuyến agent
# hiện đi theo USERNAME chứ không theo quyền, nên KHÔNG dựa vào tầng gọi để chặn — gắn cảnh báo
# ngay vào payload, cạnh chính dữ liệu, đúng cách đã chứng minh là ăn (12/08/2026).
_REPORT_NHAY_CAM = {"baocaotienluong", "baocaotongsonhansu", "baocaovipham"}
_CANH_BAO_NHAY_CAM = (
    "DỮ LIỆU NHẠY CẢM (lương / nhân sự / vi phạm). CHỈ được trả số TỔNG HỢP theo khối hoặc đơn vị. "
    "TUYỆT ĐỐI không trả lương, thu nhập hay vi phạm của một CÁ NHÂN cụ thể — câu hỏi nhắm vào một "
    "người phải từ chối và chỉ sang phòng HCNS.")


def _nhay_cam(report_type) -> bool:
    return (report_type or "") in _REPORT_NHAY_CAM


def _hop_le_so(v):
    """Phân biệt 0 THẬT / ô rỗng / lỗi công thức. Quy tất cả về 0 là cách hỏng âm thầm đã gặp
    nhiều lần: '-' và '#REF!' bị coi là 0 rồi báo 'đơn vị không phát sinh'."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return {"trang_thai": "rong"}
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("#"):
            return {"trang_thai": "loi_cong_thuc", "raw": s}
        if s in ("-", "--", "n/a", "N/A"):
            return {"trang_thai": "rong", "raw": s}
    return {"trang_thai": "co_gia_tri"}


@mcp.tool()
def source_inspect(file_name: str, sheet: str = None, max_rows: int = 40,
                   chua: str = None, quanh: int = 1) -> dict:
    """Mở file gốc (chỉ đọc) trong INPUT_DIR, Connect_VPS/received_reports, hoặc Tài liệu
    (tài liệu tham chiếu/mapping) để đào sâu số CHƯA hiển thị trên dashboard hoặc tra công thức/
    mapping. Chặn path traversal (chỉ 3 thư mục này).

    DÙNG `chua=` ĐỂ LẤY ĐÚNG DÒNG CẦN, đừng đổ cả sheet: `chua` khớp nhãn dòng (không dấu, không
    phân biệt hoa/thường) và chỉ trả các dòng khớp + `quanh` dòng lân cận mỗi bên. Đổ 200 dòng thô
    để tự dò là nguyên nhân chính làm phình ngữ cảnh (đã đo p90 = 34 KB/lượt gọi) rồi mất số vừa
    đọc khi lịch sử bị nén. Định vị trước bằng `tim_chi_tieu` thì còn nhanh hơn.

    `max_rows` mặc định 40 (trước là 200). Chỉ nâng khi thật sự cần duyệt cả sheet."""
    from .common import be_bridge as bb

    target = _resolve_readable(file_name)
    loc = bb.remove_diacritics(chua or "").strip().lower()

    wb = bb.fast_load_workbook(target, data_only=True, read_only=True)
    try:
        sheet_names = wb.sheetnames
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        rows = []
        if loc:
            tat_ca = [[("" if c is None else c) for c in r]
                      for r in ws.iter_rows(max_row=800, values_only=True)]
            khop = [i for i, r in enumerate(tat_ca)
                    if any(loc in bb.remove_diacritics(str(c)).strip().lower()
                           for c in r[:4] if c not in (None, ""))]
            giu = sorted({j for i in khop for j in range(max(0, i - quanh),
                                                         min(len(tat_ca), i + quanh + 1))})
            for j in giu[:max_rows]:
                rows.append({"dong": j + 1, "o": tat_ca[j], "khop": j in khop})
        else:
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= max_rows:
                    break
                rows.append([("" if c is None else c) for c in row])
        # Cảnh báo PHẠM VI đi kèm chính dữ liệu: model mở thẳng file theo tên (không qua
        # catalog_search) thì vẫn phải thấy "đây là số của một khối, không phải tập đoàn".
        # Xem source_catalog._canh_bao_pham_vi — đã trả nhầm lợi nhuận khối Xe tải thành lợi
        # nhuận tập đoàn 3 lần dù SKILL nêu đích danh file này (12/08/2026).
        from .common import source_catalog
        cb = None
        for e in source_catalog.search():
            if os.path.abspath(e.get("path") or "") == os.path.abspath(target):
                cb = e.get("canh_bao_pham_vi")
                break
        out = {
            "file_name": file_name, "sheet_used": ws.title, "all_sheets": sheet_names,
            "row_count_returned": len(rows), "truncated": len(rows) >= max_rows,
            "rows": rows,
        }
        if loc:
            out["loc_theo"] = {"chua": chua, "quanh": quanh}
            if not rows:
                out["canh_bao"] = (f"Không dòng nào ở 4 cột đầu chứa '{chua}' trong sheet "
                                   f"'{ws.title}'. Kiểm tra lại nhãn (dùng tim_chi_tieu) hoặc thử "
                                   f"sheet khác trong all_sheets — ĐỪNG kết luận đơn vị không có số.")
        if cb:
            out["canh_bao_pham_vi"] = cb
        for e in source_catalog.search():
            if os.path.abspath(e.get("path") or "") == os.path.abspath(target):
                if _nhay_cam(e.get("report_type")):
                    out["canh_bao_nhay_cam"] = _CANH_BAO_NHAY_CAM
                break
        return out
    finally:
        wb.close()


@mcp.tool()
def unmapped_cc_list(include_resolved: bool = False) -> dict:
    """Cost center chưa khớp MD_COSTCENTER khi điền template (admin cần bổ sung danh mục
    để lần sau tự roll-up đúng khối). Mỗi mục: raw, cong_ty, sheets, count, first/last_seen.
    Trả {"results": [...]}."""
    return {"results": memory.unmapped_cc_list(include_resolved=include_resolved)}


@mcp.tool()
def reconcile_status(dataset_id: str = None) -> dict:
    """Soi lỗ hổng pipeline: file đã kéo về chưa import (uningested_files), KPI/màn FE thiếu nguồn
    (missing_report_types/dark_kpis), và tóm tắt pipeline (collected/received/ingested). Không ghi."""
    from .common import reconcile
    return reconcile.status(dataset_id)


@mcp.tool()
def pipeline_state() -> dict:
    """View hợp nhất từng file: collected (available_metadata) -> received (indexed) -> ingested."""
    from .common import reconcile
    return reconcile.pipeline_state()


@mcp.tool()
def catalog_search(query: str = None, company: str = None, canonical_kind: str = None,
                   sheet: str = None, only_uningested: bool = False,
                   month: int = None, report_type: str = None,
                   year: int = None, ky: str = None) -> dict:
    """Tra CATALOG toàn bộ file đã kéo về (Connect_VPS/received_reports) — con trỏ lossless,
    trả lời 'có file/sheet/cột nào' tức thì (kể cả file CHƯA import). Không mở file.

    LỌC KỲ VÀ LOẠI BÁO CÁO BẰNG `month` + `report_type`, ĐỪNG nhét chúng vào `query`. Kỳ nằm
    trong tên file dưới nhiều dạng khác nhau ('M.202607', 'M202607', 'M.2026.07') nên dò bằng
    chuỗi là may rủi: hỏi query='M202607' từng chỉ ra 1/11 file tháng 7 vì 10 file kia viết
    'M.202607' — đủ để kết luận sai rằng cả tập đoàn chỉ có một báo cáo (12/08/2026).
    Đúng cách: catalog_search(report_type="baocaotaichinhrieng", month=7) -> đủ 11 đơn vị.

    DÙNG `ky='YYYY-MM'` khi có thể: `month` một mình KHÔNG định danh được kỳ — catalog đang có
    file 12/2025 nằm cạnh file 2026, nên `month=12` gom cả hai năm.

    `query` chỉ dành cho phần TÊN không đoán trước được; nhiều từ khoá thì phải khớp hết.

    Trả {"results": [...], "count": n} — mỗi mục: file, path, company, report_type, month,
    ingested, sheets:[{name,columns,nrows,canonical_kind}]. Định vị được file rồi dùng
    source_inspect đọc chi tiết ô gốc.
    """
    from .common import source_catalog
    kq = source_catalog.search(query=query, company=company, canonical_kind=canonical_kind,
                               sheet=sheet, only_uningested=only_uningested,
                               month=month, report_type=report_type, year=year, ky=ky)
    # `count` để model tự đối chiếu với kỳ vọng nghiệp vụ ("tập đoàn có 11 đơn vị") thay vì
    # đếm tay danh sách rồi kết luận thiếu/đủ.
    ra = {"results": kq, "count": len(kq)}
    if any(_nhay_cam(e.get("report_type")) for e in kq):
        ra["canh_bao_nhay_cam"] = _CANH_BAO_NHAY_CAM
    return ra


@mcp.tool()
def tim_chi_tieu(ten: str, ky: str = None, year=None, month=None, company: str = None,
                 report_type: str = None, gioi_han: int = 40) -> dict:
    """TRA VỊ TRÍ một chỉ tiêu trong 370 file / 4.838 sheet — GỌI TRƯỚC `source_inspect`.

    `catalog_search` chỉ biết TÊN file/sheet/cột, KHÔNG biết trong sheet có dòng gì, nên
    `query="doanh thu"` luôn rỗng. Tool này tra chỉ mục nhãn dòng và trả về toạ độ chính xác
    [{file, sheet, dong, ma_dong, nhan}] mà KHÔNG mở file, KHÔNG trả giá trị.

    Quy trình đúng: tim_chi_tieu(ten, ky) -> source_inspect(file, sheet, chua=<nhãn>, quanh=1).
    Mở file rồi đổ 200 dòng ra dò là cách cũ, tốn ngữ cảnh gấp hàng chục lần và hay lạc sheet.

    `ky` dạng 'YYYY-MM'. Chỉ mục dò theo NỘI DUNG THẬT nên nguồn báo cáo mới đẩy về là tìm được
    ngay, không cần ai khai báo trước."""
    from .common import row_index
    kq = row_index.tim(ten=ten, ky=ky, year=year, month=month, company=company,
                       report_type=report_type, gioi_han=gioi_han)
    if any(_nhay_cam(x.get("report_type")) for x in kq.get("ket_qua", [])):
        kq["canh_bao_nhay_cam"] = _CANH_BAO_NHAY_CAM
    return kq


@mcp.tool()
def doc_chi_tieu(chi_tieu: str = None, ky: str = None, don_vi: str = None,
                 report_type: str = "baocaotaichinhrieng", yeu_cau: list = None) -> dict:
    """ĐỌC CHỈ TIÊU CHUẨN cho MỌI đơn vị và CỘNG BẰNG CODE — dùng thay cho việc tự mở 11 file.

    Vẫn đọc thẳng Excel, không dùng DB. Trả về: `dong` (từng đơn vị, kèm sheet/dòng/mã dòng/cột đã
    dùng để kiểm ngược), `tong`, `du_lieu_du`, `don_vi_thieu`, `canh_bao`.

    **`du_lieu_du=false` thì KHÔNG được đưa số tổng** — nói rõ thiếu đơn vị nào và vì sao.
    File 'hợp nhất' đã được tách khỏi phép cộng (chống cộng đôi), xem `dong_hop_nhat`.

    THEO LÔ — dùng cho câu hỏi nhiều ý: truyền `yeu_cau=[{"y_id":"y1","chi_tieu":"doanh_thu",
    "ky":"2026-07"}, {"y_id":"y2", ...}]`. Trả một payload duy nhất kèm `da_yeu_cau`/`da_tra`/
    `con_thieu` để không rơi ý nào, và tiết kiệm ngữ cảnh so với gọi nhiều lần.

    Chỉ tiêu ngoài danh sách đã khai thì dùng `tim_chi_tieu` + `source_inspect`."""
    from .common import doc_chi_tieu as dct

    if yeu_cau:
        ket_qua, da_tra = [], []
        for i, y in enumerate(yeu_cau, 1):
            yid = (y or {}).get("y_id") or f"y{i}"
            try:
                r = dct.doc(chi_tieu=y.get("chi_tieu"), ky=y.get("ky") or ky,
                            don_vi=y.get("don_vi"), report_type=y.get("report_type") or report_type)
                r["y_id"] = yid
                if not r.get("loi"):
                    da_tra.append(yid)
            except Exception as ex:                                    # noqa: BLE001
                r = {"y_id": yid, "loi": f"{type(ex).__name__}: {ex}"[:300]}
            ket_qua.append(r)
        yeu = [(y or {}).get("y_id") or f"y{i}" for i, y in enumerate(yeu_cau, 1)]
        thieu = [x for x in yeu if x not in da_tra]
        return {"ket_qua": ket_qua, "da_yeu_cau": yeu, "da_tra": da_tra, "con_thieu": thieu,
                "bat_buoc_dau_ra": ("Mỗi y_id phải có một mục riêng trong câu trả lời. "
                                    "`con_thieu` khác rỗng thì phải nói rõ ý nào chưa trả lời được.")}

    if not chi_tieu:
        return {"loi": "Thiếu `chi_tieu` (hoặc dùng `yeu_cau` cho nhiều ý).",
                "chi_tieu_co_san": dct.danh_sach_chi_tieu()}
    return dct.doc(chi_tieu=chi_tieu, ky=ky, don_vi=don_vi, report_type=report_type)


@mcp.tool()
def danh_sach_chi_tieu_chuan() -> dict:
    """Các chỉ tiêu đã khai cho `doc_chi_tieu` + bố cục hỗ trợ. Chỉ tiêu KHÔNG có ở đây thì phải
    đi đường `tim_chi_tieu` + `source_inspect`, đừng ép vào chỉ tiêu gần giống."""
    from .common import doc_chi_tieu as dct
    return {"chi_tieu": dct.danh_sach_chi_tieu(),
            "bo_cuc": {k: v.get("_mo_ta") for k, v in dct.ban_do()["he"].items()}}


@mcp.tool()
def phan_loai_cau_hoi(cau_hoi: str) -> dict:
    """TÁCH Ý + NHẬN DIỆN LOẠI CÂU HỎI, trả về CÔNG THỨC LÀM VIỆC cho từng ý — GỌI ĐẦU TIÊN.

    Trả cho mỗi ý: loại câu hỏi, `report_type` liên quan, `cach_lay`, `bat_buoc` (ràng buộc đầu ra)
    và `bay` (các lỗi đã mắc thật với đúng loại đó), kèm kỳ/đơn vị đọc được sẵn từ câu hỏi.

    Câu hỏi nhiều ý: `y` là DANH SÁCH. Mỗi ý PHẢI có một mục riêng trong câu trả lời, đúng thứ tự
    người dùng hỏi. Ý nào không khớp loại nào vẫn nằm trong danh sách kèm `canh_bao` — phải nói rõ
    là chưa hiểu ý đó, CẤM lặng lẽ chỉ trả lời phần làm được.

    Nội dung nghiệp vụ nằm ở `loai_cau_hoi.json` — thêm nguồn báo cáo mới thì sửa JSON, không sửa
    code, không chép vào SKILL."""
    from .common import phan_loai
    return phan_loai.phan_loai(cau_hoi)


@mcp.tool()
def so_do_to_chuc(don_vi: str = None) -> dict:
    """Sơ đồ tổ chức lấy THẲNG từ master_data: công ty / khối / cost center (kèm cost center thuộc
    công ty & khối nào). Truyền `don_vi` để lọc theo mã hoặc tên.

    Dùng tool này thay vì nhớ bảng trong SKILL: cơ cấu đổi ở master_data là tool đổi theo ngay,
    còn bảng chép cứng thì lệch âm thầm."""
    from .common import be_bridge as bb
    md = bb.master_data()
    loc = bb.remove_diacritics(don_vi or "").strip().lower()

    def _khop(*vals):
        if not loc:
            return True
        return any(loc in bb.remove_diacritics(str(v or "")).strip().lower() for v in vals)

    cc = [c for c in md.get("costCenters", [])
          if _khop(c.get("ma"), c.get("ten"), c.get("congTy"), c.get("khoi"))]
    ra = {
        "cong_ty": [c for c in md.get("companies", []) if _khop(c.get("ma"), c.get("ten"))],
        "khoi": [k for k in md.get("khoi", []) if _khop(k.get("ma"), k.get("ten"))],
        "cost_center": cc,
        "nguon": "master_data (companies/khoi/costCenters)",
    }
    ra["ghi_chu"] = [
        "TC gồm 5 nhóm nội bộ, MỖI nhóm 1 file riêng: SRVF · DUAN · TRAMSAC · HO · XDV. "
        "Câu hỏi về 'TC' nói chung phải gộp cả 5, đọc 1 nhóm rồi coi là đủ là SAI.",
        "'VF'/'VinFast' KHÔNG phải mã công ty — là thương hiệu, trải trên TC + XVP + VFQN.",
        "Token thư mục hay gặp: ANTAXI/ANKHACHSAN -> AAG · GLOBALAI -> GA · "
        "XANHVINHPHUC -> XVP · HTXXANHTUYENQUANG -> HTX_XTQ · HTXXANHVINHPHUC -> HTX_XVP.",
    ]
    ra["count"] = {k: len(v) for k, v in ra.items() if isinstance(v, list)}
    return ra


@mcp.tool()
def chi_muc_trang_thai() -> dict:
    """Tình trạng chỉ mục nhãn dòng + danh sách file KHÔNG suy được kỳ.

    Chỉ mục CHƯA DỰNG khác hẳn 'không có dữ liệu' — nếu `san_sang=false` thì đừng trả lời người
    dùng là không tìm thấy số. File trong `ky_khong_ro` không xuất hiện ở mọi phép lọc theo kỳ,
    nên phải nêu ra để người ta đi sửa tên file/sidecar."""
    from .common import row_index, source_catalog
    return {"chi_muc": row_index.san_sang(), "ky_khong_ro": source_catalog.ky_khong_ro_list()}


if __name__ == "__main__":
    mcp.run()

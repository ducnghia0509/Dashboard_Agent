# -*- coding: utf-8 -*-
"""ENGINE ĐIỀN TEMPLATE VÀNG.

Nhận số liệu nguồn + mapping (cột nguồn -> cột nhập liệu của template) -> ghi vào 1
BẢN SAO Template_chuan.xlsx (giữ nguyên công thức per-row: C/D auto cty/khoi từ cost
center, các cột % tự tính). File điền xong đi qua importer_template.py sẵn có -> raw_rows.

Nguyên tắc:
- CHỈ ghi cột nhập liệu (theo tên header); KHÔNG đụng cột công thức.
- Mỗi lần điền bắt đầu từ template gốc SẠCH (rows dữ liệu trống) -> không lẫn dữ liệu cũ.
- Cost center: resolve về mã chuẩn (contract.resolve_costcenter). Không khớp -> ghi
  unmapped_cc + để trống CC -> backend suy khối = NULL -> FE gom "(Chưa phân bổ)".
"""
import io
import os


from .common import be_bridge as bb
from .common import contract as C
from .common import memory
from .common import source_catalog as SC

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
FILLED_DIR = os.path.normpath(os.path.join(_AGENT_ROOT, "..", "template_trust", "filled"))


def _norm(s) -> str:
    # Bỏ marker UI "◀ NHẬP" (đánh dấu cột nhập trên template) — mapping học đôi khi kèm marker này
    # trong khi header cột đích KHÔNG có (hoặc ngược lại) -> phải chuẩn hoá để khớp, nếu không cột
    # đối tượng (Khách hàng/Nhà cung cấp/Loại TS...) sẽ KHÔNG được ghi -> file điền thiếu cột.
    t = "" if s is None else str(s)
    t = t.split("◀")[0]   # cắt phần marker "◀ NHẬP" trở đi
    return bb.remove_diacritics(t).strip().lower()


def _col_index(sheet_spec) -> dict:
    """{norm(header) -> 1-based column} cho sheet đích."""
    return {_norm(h): j + 1 for j, h in enumerate(sheet_spec["columns"]) if h}


def _is_costcenter_col(header: str) -> bool:
    n = _norm(header)
    return "cost center" in n or "costcenter" in n


def fill(target_sheet: str, records: list, out_path: str, source_template: str = None) -> dict:
    """Ghi list record (dict {template_header: value}) vào bản sao template.

    records: mỗi phần tử là 1 dòng nhập liệu; key = TÊN CỘT template (input), value = số/chuỗi.
    CHỈ ghi cột có trong records (bỏ qua cột công thức). Trả {rows_written, out_path}.
    """
    src = source_template or C.GOLDEN_TEMPLATE
    spec = C.data_sheets().get(target_sheet)
    if spec is None:
        raise ValueError(f"Sheet đích không hợp lệ: {target_sheet}. Hợp lệ: {list(C.data_sheets())}")
    cidx = _col_index(spec)
    start_row = spec["header_row"] + 2  # 1-based, ngay sau dòng header

    wb = bb.fast_load_workbook(src)  # data_only=False -> giữ công thức
    try:
        ws = wb[target_sheet]
        written, cells = 0, 0
        dropped = set()
        for i, rec in enumerate(records):
            r = start_row + i
            for key, val in rec.items():
                if val is None or val == "":
                    continue
                ci = cidx.get(_norm(key))
                if ci:  # chỉ ghi cột nhập liệu khớp header; cột công thức bị bỏ qua
                    ws.cell(row=r, column=ci, value=val)
                    cells += 1
                else:
                    # key không khớp cột template — KHÔNG nuốt im lặng (rows_written 100
                    # mà file chỉ có cột Kỳ): báo lại để caller thấy dữ liệu bị rơi.
                    dropped.add(str(key))
            written += 1
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        wb.save(out_path)
        out = {"rows_written": written, "cells_written": cells,
               "out_path": out_path, "target_sheet": target_sheet}
        if dropped:
            out["dropped_keys"] = sorted(dropped)
        return out
    finally:
        wb.close()


def _read_source(data: bytes, sheet: str):
    """Đọc sheet nguồn -> (header:list[str], rows:list[list]) (data_only)."""
    wb = bb.fast_load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return rows


def _all_sheet_headers(data: bytes, n: int = 20) -> dict:
    """Mở workbook 1 LẦN, trả {sheet: [n dòng đầu]} — đủ để dò header/fingerprint.
    QUAN TRỌNG: file BCTC nặng (vd HO 8MB, sheet CĐPS 16350 cột) tốn ~18s/lần load;
    mở lại theo TỪNG sheet sẽ treo. Load 1 lần + đọc header mọi sheet chỉ ~0.03s."""
    wb = bb.fast_load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        out = {}
        for ws in wb.worksheets:
            rows = []
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i >= n:
                    break
                rows.append(list(r))
            out[ws.title] = rows
        return out
    finally:
        wb.close()


def _is_money_col(header: str) -> bool:
    """Cột tiền của template (đơn vị tỷ) — header chứa '(tỷ)' hoặc '(tỷ/xe)', '(tỷ/...)'."""
    return "(ty" in _norm(header)


# Ngưỡng biên độ cột tiền (đơn vị TỶ). Doanh thu tập đoàn cỡ vài nghìn tỷ; > 1 triệu tỷ gần như
# chắc chắn là VND CHƯA quy đổi (quên value_scale=1e-9) -> chặn để không lạm phát 1e9 lần (B9).
_MONEY_SANITY_MAX_TY = 1_000_000.0


def _max_money_abs(records: list) -> float:
    """|giá trị| lớn nhất trong các cột tiền '(tỷ)' của records (0.0 nếu không có)."""
    mx = 0.0
    for rec in records:
        for k, v in rec.items():
            if isinstance(v, (int, float)) and _is_money_col(k):
                mx = max(mx, abs(v))
    return mx


def _mapping_spec(src_spec):
    """Giá trị mapping: 'TÊN CỘT NGUỒN' (str) hoặc dict pivot bảng dọc:
    {'src': 'CỘT GIÁ TRỊ', 'khi': {CỘT ĐIỀU KIỆN: đk}, 'ffill': bool} — đk: chuỗi (khớp
    chứa-trong, không dấu, trên giá trị FORWARD-FILL để chịu được cột section chỉ ghi 1 lần
    rồi merge/để trống), '' (ô RAW phải RỖNG — lọc dòng subtotal), '*' (ô RAW phải KHÁC rỗng).
    'ffill': True -> GIÁ TRỊ lấy theo forward-fill (vd cột CÔNG TY chỉ ghi ở dòng nhóm, các
    dòng ngân hàng bên dưới để trống nhưng vẫn thuộc công ty đó).
    Trả (src_header, khi_dict|None, ffill_bool)."""
    if isinstance(src_spec, dict):
        return (src_spec.get("src") or src_spec.get("source"),
                (src_spec.get("khi") or src_spec.get("when") or {}),
                bool(src_spec.get("ffill")))
    return src_spec, None, False


def build_records(source_rows: list, header_row: int, mapping: dict,
                   target_sheet: str, period: str = None, file_name: str = None,
                   value_scale: float = 1.0, constants: dict = None) -> dict:
    """Từ dữ liệu nguồn + mapping -> list record cho template + resolve cost center.

    mapping: {template_header: source_header | {'src':..., 'khi': {...}}} — dạng dict cho phép
      PIVOT bảng dọc (vd TC01_SD TIỀN: mỗi dòng = LOẠI TIỀN×công ty×ngân hàng) thành cột ngang
      của template (Tiền mặt/Tiền gửi/Tiền vay) bằng điều kiện dòng — xem _mapping_spec.
    constants: {template_header: giá trị HẰNG} — gán cứng cho mọi dòng (vd cost center/công ty khi
      sheet nguồn cấp công ty, không có CC theo dòng). Áp SAU mapping, TRƯỚC resolve CC.
    period: điền cột 'Kỳ' nếu template có và mapping không có.
    value_scale: nhân giá trị SỐ ở cột tiền '(tỷ)' (vd nguồn VND -> tỷ dùng 1e-9). Cột text không đổi.
    Trả {records, unresolved:[...], resolved_cc:int, header_warnings:[...]}.
    """
    constants = constants or {}
    hdr = [ _norm(h) for h in source_rows[header_row] ]
    # Header TRÙNG TÊN (điển hình sheet 131/331: Nợ/Có lặp 3 lần dưới Đầu kỳ/Phát sinh/
    # Cuối kỳ): dict-comprehension cũ giữ cột CUỐI một cách âm thầm -> map nhầm chiều số
    # liệu. Giữ cột ĐẦU (trái nhất, dễ đoán hơn) và BÁO warning để analyst/orchestrator
    # biết mapping theo tên cột là mơ hồ, phải dùng mapping theo chỉ số cột thay thế.
    src_idx, _dup_hdr = {}, set()
    for j, h in enumerate(hdr):
        if not h:
            continue
        if h in src_idx:
            _dup_hdr.add(h)
        else:
            src_idx[h] = j
    _used_src = {_norm(_mapping_spec(v)[0]) for v in mapping.values() if _mapping_spec(v)[0]}  # noqa: E501
    header_warnings = [
        f"Header nguồn trùng tên '{h}' (nhiều cột cùng nhãn) — mapping theo tên chỉ lấy "
        f"cột trái nhất, các cột trùng còn lại KHÔNG được đọc."
        for h in sorted(_dup_hdr) if h in _used_src
    ]
    # cột template nào là cost center (để resolve)
    spec = C.data_sheets()[target_sheet]
    cc_headers = [c for c in spec["columns"] if _is_costcenter_col(c)]
    period_headers = [c for c in spec["columns"] if _norm(c).startswith("ky") or _norm(c).startswith("ngay")]

    records, unresolved = [], []
    resolved_cc = 0
    _ffill: dict = {}   # cột điều kiện -> giá trị non-empty gần nhất (forward-fill cho cột section)
    for r in source_rows[header_row + 1:]:
        if not r or all(c in (None, "") for c in r):
            continue
        # cập nhật forward-fill cho MỌI cột nguồn (rẻ, chỉ cột có header)
        for h, j in src_idx.items():
            if j < len(r) and r[j] not in (None, "") and str(r[j]).strip() != "":
                _ffill[h] = r[j]
        rec = {}
        for tmpl_h, src_spec in mapping.items():
            src_h, khi, use_ffill = _mapping_spec(src_spec)
            j = src_idx.get(_norm(src_h)) if src_h else None
            if j is not None and j < len(r):
                if khi:
                    ok = True
                    for cond_col, cond in khi.items():
                        cj = src_idx.get(_norm(cond_col))
                        raw_cell = r[cj] if cj is not None and cj < len(r) else None
                        raw_empty = raw_cell in (None, "") or str(raw_cell).strip() == ""
                        if cond == "":       # ô RAW phải RỖNG (vd NGÂN HÀNG='' -> chỉ lấy dòng công ty)
                            ok = raw_empty
                        elif cond == "*":    # ô RAW phải KHÁC rỗng (vd CÔNG TY='*' -> bỏ dòng tổng section)
                            ok = not raw_empty
                        else:                # khớp chứa-trong (không dấu) trên giá trị forward-fill
                            ok = _norm(cond) in _norm(_ffill.get(_norm(cond_col)))
                        if not ok:
                            break
                    if not ok:
                        continue
                val = _ffill.get(_norm(src_h)) if use_ffill else r[j]
                if _is_money_col(tmpl_h):
                    # Nguồn hay để số dạng TEXT ("1.000.000.000"): nếu không parse về số
                    # trước thì chuỗi VND vừa KHÔNG được nhân value_scale vừa lọt qua
                    # _max_money_abs (chỉ soi int/float) -> qua mặt guard B9 rồi
                    # importer parse thành giá trị phồng 1e9 lần. Parse được -> dùng số;
                    # không parse được (text thật) -> giữ nguyên.
                    if isinstance(val, str) and val.strip():
                        num = bb.parse_num(val)
                        if num or val.strip() in ("0", "-", "–"):
                            val = num
                    if value_scale != 1.0 and isinstance(val, (int, float)):
                        val = val * value_scale   # VND -> tỷ (giữ cột text nguyên vẹn)
                rec[tmpl_h] = val
        # Dòng nguồn KHÔNG cho ra giá trị mapping nào (vd bị mọi 'khi' filter loại, hoặc dòng
        # trang trí) -> BỎ, không tạo record chỉ có Kỳ/constants — chính loại dòng "chỉ có cột
        # Kỳ" từng làm importer báo 'chưa có dòng dữ liệu' rất khó truy (03B, 2026-07-10).
        if not any(v not in (None, "") for v in rec.values()):
            continue
        for tmpl_h, cval in constants.items():   # hằng số (CC/công ty) gán cứng mọi dòng
            rec[tmpl_h] = cval
        if period and period_headers:
            rec.setdefault(period_headers[0], period)
        # resolve cost center trong record (nếu có cột CC được map)
        for cc_h in cc_headers:
            raw_cc = rec.get(cc_h)
            if raw_cc in (None, ""):
                continue
            res = C.resolve_costcenter(raw_cc)
            # FUZZY = KHÔNG chắc chắn -> KHÔNG tự canonical hoá (nguyên tắc "không đoán"): coi như chưa
            # map, ghi unmapped để analyst xác nhận, giữ nguyên giá trị raw (B13). Chỉ nhận exact/alias.
            if res and not str(res.get("matched_by", "")).startswith("fuzzy"):
                rec[cc_h] = res["cc"]        # chuẩn hoá về mã CC chuẩn
                resolved_cc += 1
            else:
                memory.unmapped_cc_record(str(raw_cc), sheet=target_sheet, file_name=file_name)
                unresolved.append(str(raw_cc))
                # KHÔNG khớp cost_centers.yaml -> để TRỐNG (đúng docstring dòng 11-12), TUYỆT ĐỐI
                # không giữ mã do analyst tự bịa trôi vào raw_rows (vụ CP_AAG/XT_BP, 2026-07-09).
                rec[cc_h] = ""
        if any(v not in (None, "") for v in rec.values()):
            records.append(rec)
    return {"records": records, "unresolved": sorted(set(unresolved)), "resolved_cc": resolved_cc,
            "header_warnings": header_warnings}


def import_filled(path: str, cong_ty: str = None, khoi: str = None, source_file: str = None) -> dict:
    """Nạp 1 file ĐÃ ĐIỀN (template chuẩn) vào raw_rows.

    GỘP ĐA-CÔNG-TY: mỗi KỲ (tháng) = 1 dataset dùng CHUNG cho mọi công ty; import chỉ thay
    dữ liệu của (kỳ, cong_ty) NÀY (không xoá công ty khác) rồi ĐÓNG DẤU cong_ty cho các dòng
    vừa nạp -> dashboard tập đoàn thấy đủ công ty, filter theo công ty hoạt động.
    (Không truyền cong_ty -> quay lại hành vi cũ: thay cả kỳ.)"""
    with open(path, "rb") as fh:
        data = fh.read()
    _wb = bb.fast_load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        is_tmpl = bb.detect_template(_wb)
    finally:
        _wb.close()
    if not is_tmpl:
        return {"ok": False, "error": "File điền không nhận diện là template chuẩn (thiếu sheet 01..12?)."}
    parsed = bb.template_parse(data)
    if not parsed["tuples"]:
        return {"ok": False, "error": "Template chưa có dòng dữ liệu nào (kiểm tra cột Kỳ)."}
    grain = parsed["grain"]
    period = parsed.get("period")
    db = bb.get_db()

    # CÔNG TY chỉ nhận MÃ HỢP LỆ trong MD_CONGTY. Ưu tiên TOKEN THƯ MỤC NGUỒN (phần trước '::' của
    # source_file, vd 'HTXXANHVINHPHUC') qua _COMPANY_FOLDER_ALIAS — tín hiệu định danh MẠNH NHẤT,
    # tách được pháp nhân dùng-chung-nhãn trong tên file (3 Taxi Xanh đều 'B.6.XVP'). Chỉ thư mục
    # ĐƠN-pháp-nhân có trong alias mới resolve; thư mục đa pháp nhân (THUCHI/SRVF) -> None -> giữ tag
    # per-dòng theo cost center. Fallback: suy từ raw cong_ty + tên file 'B.<n>.<MÃ>.' (hành vi cũ).
    _by_folder = C.resolve_company(source_file.split("::", 1)[0]) if (source_file and "::" in source_file) else None
    cong_ty = _by_folder or C.resolve_company(cong_ty, os.path.basename(path))

    file_types = sorted({t[0] for t in parsed["tuples"] if t[0]})
    # Xác định / TÁI DÙNG dataset đích TRƯỚC — KHÔNG xoá gì ở bước này (atomic swap: xoá bản cũ chỉ
    # sau khi insert mới thành công, xem dưới). Trước đây delete chạy TRƯỚC insert (cả nhánh month
    # lẫn day) -> insert lỗi giữa chừng thì mất luôn dữ liệu cũ, cùng lớp lỗi đã sửa ở main.py (B1).
    prev_day_ids = []
    if grain == "month" and period:
        # 1 dataset dùng chung/kỳ. TÁI DÙNG (không tạo trùng), kích hoạt để dashboard hiện.
        existing = [d for d in bb.repo.list_datasets("month") if d.get("period") == period]
        target = existing[0]["id"] if existing else bb.repo.create_dataset(period, kind="month", period=period)["id"]
        bb.repo.set_active(target)
    elif grain == "day" and parsed["ngay"]:
        name = f"Ngày {parsed['ngay']}"
        prev_day_ids = [d["id"] for d in bb.repo.list_datasets("day") if d.get("name") == name]
        target = bb.new_dataset(kind="day", period=period, name=name)["id"]
    else:
        target = bb.new_dataset(kind=grain, period=period, name=None)["id"]

    # GUARD CHÉO-FILE (chống nạp đôi khi vai kqkd/bctc được nới bằng nhau, 2026-07-10):
    # BaocaoHQKD và Baocaotaichinhrieng là 2 bản export CÙNG sổ — user chỉ chạy MỘT trong hai.
    # Nếu report_type của file này ĐÃ có dòng từ FILE KHÁC cùng dataset+công ty -> CHẶN kèm tên
    # file kia (muốn thay nguồn thì xoá/hide file cũ trước). Chỉ áp khi biết cả source_file lẫn
    # cong_ty (đường analyst/template_fill); extract_thuchi/sodutien (cong_ty=None) tự idempotent
    # theo file, không đi qua guard này.
    if grain == "month" and file_types and source_file and cong_ty:
        tph = ",".join(["?"] * len(file_types))
        rc = db.execute(
            f"SELECT DISTINCT source_file FROM raw_rows WHERE dataset_id=? AND cong_ty=? "
            f"AND report_type IN ({tph}) AND source_file IS NOT NULL AND source_file<>?",
            [target, cong_ty, *file_types, source_file]).fetchall()
        # CHỈ CHẶN khi nguồn kia CÙNG GỐC THẬT (cùng tiền tố '<công_ty_thư_mục>::') — đó mới là 2
        # bản export cùng 1 sổ của CÙNG 1 nguồn (BaocaoHQKD vs Baocaotaichinhrieng). Nguồn KHÁC THẬT
        # (khác thư mục nhận, vd 3 HTX cùng nộp 'B.6.XVP...') tuy chia sẻ mã pháp nhân (XVP) nhưng là
        # dữ liệu ĐỘC LẬP -> CHO PHÉP cùng tồn tại (tách nguồn triệt để).
        def _src_root(s):
            return s.split("::", 1)[0] if "::" in (s or "") else None
        my_root = _src_root(source_file)
        others = sorted({dict(r)["source_file"] for r in rc
                         if _src_root(dict(r)["source_file"]) == my_root})
        if others:
            return {"ok": False, "error": (
                f"CHẶN NẠP TRÙNG: report_type {sorted(file_types)} của kỳ {period} (công ty "
                f"{cong_ty}) đã được nạp từ file khác: {others}. BaocaoHQKD và "
                f"Baocaotaichinhrieng là 2 bản export cùng sổ — chỉ chạy MỘT file/kỳ. "
                f"Muốn thay nguồn: xoá/hide dữ liệu file cũ trước rồi nạp lại.")}

    before_id = db.execute("SELECT COALESCE(MAX(id),0) m FROM raw_rows").fetchone()["m"]
    result = bb.template_import_parsed(target, parsed, source_file=source_file)  # INSERT TRƯỚC

    # XOÁ bản CŨ — CHỈ SAU KHI insert mới đã thành công (atomic swap thật sự).
    if grain == "month" and file_types:
        # IDEMPOTENT THEO NGUỒN: source_file định danh nguồn DUY NHẤT (folder::file) -> mỗi file chỉ
        # thay ĐÚNG dòng của CHÍNH NÓ; 2 file khác nhau (kể cả TRÙNG BASENAME giữa 2 công ty) KHÔNG đè
        # nhau (fix B2/B3). KHÔNG wipe-toàn-kỳ: thiếu cả source_file lẫn cong_ty -> KHÔNG xoá.
        # QUAN TRỌNG (fix 2026-07-16): khi CÓ source_file thì xoá theo source_file THÔI, KHÔNG kèm
        # cong_ty. Trước đây AND cả cong_ty -> khi cong_ty của lần nạp ĐỔI so với bản cũ (điển hình
        # '' -> 'AAG'/'TC' sau khi bật resolve theo thư mục) thì bản cũ (cong_ty khác) KHÔNG bị xoá ->
        # TRÙNG BẢN (đã gặp: An Taxi TSCĐ, DUAN T01). source_file đã đủ định danh -> mọi dòng cùng
        # source_file+report_type là của CHÍNH file này, thay sạch bất kể cong_ty. Chỉ khi KHÔNG có
        # source_file (đường cũ) mới scope theo cong_ty.
        if source_file or cong_ty:
            tph = ",".join(["?"] * len(file_types))
            conds = ["dataset_id=?", f"report_type IN ({tph})", "id<=?"]
            params = [target, *file_types, before_id]
            if source_file:
                conds.append("source_file=?"); params.append(source_file)
            else:
                conds.append("cong_ty=?"); params.append(cong_ty)
            db.execute("DELETE FROM raw_rows WHERE " + " AND ".join(conds), params)
            db.commit()
        else:
            print(f"[import_filled] WARN: bỏ qua xoá bản cũ vì thiếu cả source_file lẫn cong_ty "
                  f"(kỳ={period}) — tránh wipe toàn kỳ; có thể phát sinh dòng trùng.", flush=True)
    elif grain == "day" and prev_day_ids:
        # Giữ nguyên ngữ nghĩa cũ (1 dataset/ngày, thay cả dataset) nhưng AN TOÀN: dataset mới đã
        # insert xong mới xoá dataset cũ, không còn khoảng hở mất dữ liệu khi insert lỗi.
        for old_id in prev_day_ids:
            bb.repo.delete_dataset(old_id)

    # ĐÓNG DẤU công ty CHỈ cho dòng CHƯA suy được công ty từ CC (fallback) -> KHÔNG ghi đè
    # công ty đã resolve đúng từ cost center (tránh folder-token 'ANTAXI' đè công ty thật 'AAG').
    if cong_ty:
        db.execute("UPDATE raw_rows SET cong_ty=? WHERE dataset_id=? AND id>? "
                   "AND (cong_ty IS NULL OR cong_ty='')", [cong_ty, target, before_id])
        db.commit()
    # ĐÓNG DẤU KHỐI (HIỆU QUẢ KD gán theo Khối suy từ đường dẫn nguồn) CHỈ cho dòng CHƯA có khối
    # (khối suy từ cost center vẫn giữ nguyên). KHÔNG stamp cho dòng tiền THUCHI (theo pháp nhân).
    if khoi:
        # canonical hoá về ĐÚNG tên khối metric dùng (master.khoi_names) — khớp hoa/thường (Template
        # 'Xe điện' vs master 'xe điện') để không tạo card khối "lạ" tách khỏi danh mục.
        canon = {n.strip().upper(): n for n in bb.khoi_names()}
        khoi = canon.get(str(khoi).strip().upper(), khoi)
        db.execute("UPDATE raw_rows SET khoi=? WHERE dataset_id=? AND id>? "
                   "AND (khoi IS NULL OR khoi='') AND report_type<>'THUCHI'", [khoi, target, before_id])
        db.commit()
    if grain == "month" and result.get("period"):
        bb.set_period(target, result["period"])
    return {"ok": True, "dataset_id": target, "grain": grain, "cong_ty": cong_ty,
            "rows_imported": result["rows"], "by_type": result["by_type"],
            "period": result.get("period"), "ngay": result.get("ngay")}


def _derive_kqkd_summary(records: list, columns: list) -> list:
    """Sinh 2 dòng chuẩn cho 01_HQKD để KPI Chi phí & LNTT sáng (nguồn KQKD liệt kê chi phí rời,
    không có dòng 'Tổng chi phí'; dòng LNTT tên 'Tổng lợi nhuận... trước thuế' không khớp code).

    - Tổng chi phí = Σ Thực hiện các dòng chi phí (giá vốn/CPTC/CPBH/CPQLDN/CP khác).
    - Lợi nhuận trước thuế = Thực hiện dòng chứa 'loi nhuan'+'truoc thue'.
    Nhóm theo (Kỳ, Cost center). KHÔNG thêm nếu đã có sẵn dòng chuẩn (tránh trùng)."""
    chi = next((h for h in columns if "chi tieu" in _norm(h)), None)
    th = next((h for h in columns if "thuc hien" in _norm(h)), None)
    if not chi or not th:
        return records
    ky = next((h for h in columns if _norm(h).startswith(("ky", "ngay"))), None)
    cc = next((h for h in columns if _is_costcenter_col(h)), None)
    from collections import defaultdict
    G = defaultdict(lambda: {"cost": 0.0, "lntt": None, "lntt_len": None,
                             "has_cp": False, "has_lntt": False, "proto": None})
    for r in records:
        ct = _norm(r.get(chi))
        key = (r.get(ky) if ky else None, r.get(cc) if cc else None)
        g = G[key]
        if g["proto"] is None:
            g["proto"] = r
        if ct == "tong chi phi":
            g["has_cp"] = True
        if ct == "loi nhuan truoc thue":
            g["has_lntt"] = True
        v = r.get(th)
        if isinstance(v, (int, float)):
            # BỎ QUA dòng con "Trong đó: ..." VÀ mọi dòng 'lãi vay' (lãi vay ⊂ chi phí tài chính)
            # -> chỉ cộng dòng NHÓM cấp cao, tránh đếm trùng lãi vay trong Tổng chi phí (B18).
            is_sub = "trong do" in ct or "lai vay" in ct
            if not is_sub and any(p in ct for p in C.KQKD_COST_PATTERNS):
                g["cost"] += v
            # LNTT: loại EBIT / 'trước thuế và lãi vay'; chọn dòng tên NGẮN NHẤT (tổng cấp cao),
            # KHÔNG last-wins (tránh bắt nhầm EBIT xuất hiện sau) (B18).
            if (all(p in ct for p in C.KQKD_LNTT_REQUIRE)
                    and "ebit" not in ct and "lai vay" not in ct
                    and (g["lntt"] is None or len(ct) < g["lntt_len"])):
                g["lntt"] = v
                g["lntt_len"] = len(ct)
    derived = []
    for (ky_v, cc_v), g in G.items():
        def _mk(name, val):
            rec = {chi: name, th: val}
            if ky:
                rec[ky] = ky_v
            if cc and cc_v is not None:
                rec[cc] = cc_v
            return rec
        if g["cost"] and not g["has_cp"]:
            derived.append(_mk("Tổng chi phí", round(g["cost"], 9)))
        if g["lntt"] is not None and not g["has_lntt"]:
            derived.append(_mk("Lợi nhuận trước thuế", round(g["lntt"], 9)))
    return records + derived


def _header_row_of(rows) -> int:
    """Dòng header = dòng CÓ NHIỀU Ô NHẤT trong 20 dòng đầu (dòng đầu tiên đạt max).
    Bền với báo cáo tài chính thật có nhiều dòng tiêu đề/metadata (Đơn vị/Mẫu số...) ở trên."""
    best_i, best_n = 0, -1
    for i, r in enumerate(rows[:20]):
        n = sum(1 for c in r if c not in (None, ""))
        if n > best_n:
            best_n, best_i = n, i
    return best_i


def _apply_rename_rows(records: list, target_sheet: str, rename_rows: dict) -> int:
    """ÁNH XẠ TÊN CHỈ TIÊU về CHUẨN cho ĐÚNG dòng analyst chỉ định (row-level, khác mapping cột).
    rename_rows = {tên chỉ tiêu NGUỒN (khớp không dấu, chứa-trong): tên chuẩn}. Dùng cho 01_HQKD:
    analyst chọn ĐÚNG 1 dòng tổng doanh thu -> 'Doanh thu thuần', 1 dòng tổng chi phí -> 'Tổng chi phí',
    1 dòng LNTT -> 'Lợi nhuận trước thuế' (chọn 1 dòng/chỉ tiêu để KHỎI cộng trùng). Trả số dòng đã đổi."""
    if not rename_rows:
        return 0
    chi = next((h for h in (records[0].keys() if records else []) if "chi tieu" in _norm(h)), None)
    if not chi:
        return 0
    n = 0
    for src, canonical in rename_rows.items():
        if not src or not canonical:
            continue
        sn = _norm(src)
        if not sn:
            continue
        # Mỗi key đổi ĐÚNG 1 dòng: ưu tiên khớp CHÍNH XÁC; nếu không có, lấy dòng chứa-trong NGẮN
        # NHẤT (tổng cấp cao nhất, tránh khớp nhầm dòng con dài hơn) -> KHÔNG cộng trùng.
        exact = [r for r in records if _norm(r.get(chi)) == sn]
        cands = exact or [r for r in records if sn in _norm(r.get(chi))]
        if not cands:
            continue
        best = min(cands, key=lambda r: len(_norm(r.get(chi))))
        if best.get(chi) != canonical:
            best[chi] = canonical
            n += 1
    return n


def fill_from_source(data: bytes, source_sheet: str, target_sheet: str, mapping: dict,
                     period: str = None, cong_ty: str = None, file_name: str = None,
                     dry_run: bool = True, auto_import: bool = False, learn: bool = True,
                     value_scale: float = 1.0, constants: dict = None,
                     normalize_kqkd: bool = True, rename_rows: dict = None,
                     khoi: str = None, source_path: str = None) -> dict:
    """Pipeline: đọc nguồn -> build records (+resolve CC) -> (dry_run: preview | ghi file điền
    [+ auto_import: nạp raw_rows]). value_scale nhân cột tiền (VND->tỷ dùng 1e-9). constants gán cứng
    (vd cost center). rename_rows: đổi tên CHỈ TIÊU của đúng dòng nguồn về tên chuẩn (xem _apply_rename_rows
    — cần cho 01_HQKD để KPI Doanh thu/LNTT sáng). normalize_kqkd: với 01_HQKD tự sinh dòng 'Tổng chi phí'
    & 'Lợi nhuận trước thuế' NẾU nguồn chưa có. learn=True: khi ghi thật, LƯU mapping+scale+constants."""
    rows = _read_source(data, source_sheet)
    if not rows:
        return {"ok": False, "error": f"Sheet nguồn rỗng: {source_sheet}"}
    # FAIL-LOUD khi key template-side của mapping/constants KHÔNG khớp cột template: trước
    # đây fill() lặng lẽ bỏ qua key lạ -> file điền chỉ có cột Kỳ, import báo "chưa có dòng
    # dữ liệu (kiểm tra cột Kỳ)" rất khó truy (vụ 03B_SODU_TIEN 2026-07-10: mapping dùng
    # "Tổng số dư" thay vì header thật "Tổng số dư tiền (tỷ)"). Chặn ngay + kê cột hợp lệ
    # để analyst tự sửa, không để rơi xuống lỗi mờ ở bước import.
    _spec = C.data_sheets().get(target_sheet)
    if _spec is None:
        return {"ok": False, "error": f"Sheet đích không hợp lệ: {target_sheet}. "
                                      f"Hợp lệ: {list(C.data_sheets())}"}
    # PHÂN VAI FILE NGUỒN (ma trận Book1 — xem contract.ROLE_TARGETS): file thu chi chỉ được
    # cấp dòng tiền; file KQKD/BCTC (2 bản export CÙNG sổ) mỗi bên chỉ cấp nhóm sheet của vai
    # mình -> chặn nạp ĐÔI cùng dữ liệu dưới 2 source_file (đếm x2) và LCTT lọt vào dòng tiền.
    if file_name and not C.role_allows(file_name, target_sheet):
        _role = C.file_role(file_name)
        return {"ok": False, "error": (
            f"File '{file_name}' có vai '{_role}' — KHÔNG được nạp vào {target_sheet}. "
            f"Vai này chỉ được nạp: {sorted(C.ROLE_TARGETS[_role])}. Sheet này lấy từ loại "
            f"file khác (xem ma trận nguồn dữ liệu), bỏ qua để tránh nạp trùng đôi.")}
    _valid = set(_col_index(_spec))
    _bad = [k for k in list(mapping or {}) + list((constants or {}))
            if _norm(k) not in _valid]
    if _bad:
        return {"ok": False, "error": (
            f"Mapping/constants có key KHÔNG khớp cột template của {target_sheet}: {_bad}. "
            f"Phải dùng ĐÚNG tên header (tra template_contract_info). "
            f"Cột hợp lệ: {_spec['columns']}")}
    header_row = _header_row_of(rows)
    src_cols = [c for c in rows[header_row]] if header_row < len(rows) else []
    fp = memory.source_fingerprint(source_sheet, src_cols)
    built = build_records(rows, header_row, mapping, target_sheet, period=period,
                          file_name=file_name, value_scale=value_scale, constants=constants)
    renamed_n = _apply_rename_rows(built["records"], target_sheet, rename_rows)  # TRƯỚC derive
    # Sheet đích grain THÁNG: ÉP cột Kỳ = kỳ tháng (đè nếu mapping lỡ ghi NGÀY vào Kỳ) -> tránh
    # grain bị lật thành 'day' rồi import vào dataset day (bị delete_by_key day đè, mất dữ liệu).
    if period and C.grain_for(target_sheet) == "month":
        chi_ky = next((h for h in (built["records"][0].keys() if built["records"] else [])
                       if _norm(h).startswith("ky")), None)
        if chi_ky:
            for r in built["records"]:
                r[chi_ky] = period
    derived_n = 0
    if normalize_kqkd and target_sheet == "01_HQKD":
        cols = C.data_sheets().get("01_HQKD", {}).get("columns", [])
        before = len(built["records"])
        built["records"] = _derive_kqkd_summary(built["records"], cols)
        derived_n = len(built["records"]) - before
    # GUARD BIÊN ĐỘ (B9): giá trị cột tiền quá lớn -> gần như chắc quên value_scale=1e-9 (VND chưa đổi).
    max_money = _max_money_abs(built["records"])
    scale_warn = max_money > _MONEY_SANITY_MAX_TY
    result = {
        "ok": True, "target_sheet": target_sheet, "grain": C.grain_for(target_sheet),
        "source_fingerprint": fp,
        "row_count": len(built["records"]),
        "renamed_rows": renamed_n,  # số dòng đã đổi tên về chuẩn (rename_rows)
        "derived_kqkd": derived_n,  # số dòng chuẩn tự sinh (Tổng chi phí / LNTT)
        "resolved_cc": built["resolved_cc"], "unresolved_cc": built["unresolved"],
        # sample: kèm các dòng suy diễn ở cuối để kiểm
        "sample": (built["records"][:4] + built["records"][-derived_n:]) if derived_n else built["records"][:5],
        "dry_run": dry_run,
        "max_money_ty": round(max_money, 4),
        "scale_warning": scale_warn,
        "header_warnings": built.get("header_warnings") or [],
    }
    if scale_warn:
        result["error"] = (
            f"BIÊN ĐỘ BẤT THƯỜNG: giá trị cột tiền lớn nhất = {max_money:,.0f} tỷ "
            f"(> {_MONEY_SANITY_MAX_TY:,.0f}). Nhiều khả năng nguồn là VND và QUÊN value_scale=1e-9. "
            f"Kiểm tra lại; nếu nguồn đúng là VND hãy truyền value_scale=1e-9.")
    if dry_run:
        return result
    if scale_warn:
        # KHÔNG ghi file / KHÔNG nạp raw_rows / KHÔNG học spec khi biên độ bất thường (B9).
        result["ok"] = False
        return result
    tag = "_".join(x for x in [cong_ty, period, target_sheet] if x) or target_sheet
    out_path = os.path.join(FILLED_DIR, f"{tag}.xlsx".replace("/", "-"))
    w = fill(target_sheet, built["records"], out_path)
    result["out_path"] = w["out_path"]
    result["rows_written"] = w["rows_written"]
    if learn and built["records"]:
        # constants đặc thù công ty (vd cost center) -> lưu spec theo SCOPE công ty để KHÔNG replay
        # nhầm sang công ty khác cùng layout (B11). Layout KHÔNG constants -> lưu chung (tái dùng đa cty).
        if constants:
            comp = C.resolve_company(cong_ty, file_name)
            if comp:
                save_fp = memory.source_fingerprint(source_sheet, src_cols, scope=comp)
                memory.fill_spec_save(save_fp, target_sheet, mapping, source_sheet=source_sheet,
                                      value_scale=value_scale, constants=constants, rename_rows=rename_rows)
                result["learned"] = True
                result["source_fingerprint"] = save_fp
            else:
                # KHÔNG suy được công ty -> KHÔNG lưu spec mang constants (nếu lưu unscoped, autofill
                # sau này có thể áp nhầm constants — vd cost center — của công ty này cho công ty khác
                # cùng layout). Thà chưa tự động hoá được còn hơn học sai (B11).
                result["learned"] = False
                result["learn_skipped_reason"] = (
                    "constants có dữ liệu đặc thù công ty nhưng không suy được mã công ty hợp lệ "
                    "-> không lưu spec để tránh lây nhiễm chéo sang công ty khác cùng layout.")
        else:
            memory.fill_spec_save(fp, target_sheet, mapping, source_sheet=source_sheet,
                                  value_scale=value_scale, constants=constants, rename_rows=rename_rows)
            result["learned"] = True
            result["source_fingerprint"] = fp
    if auto_import:
        # resolve công ty HỢP LỆ từ TÊN FILE NGUỒN (import_filled chỉ thấy tên file đã điền,
        # không có mã 'B.7.AAG' để suy) -> ANTAXI/DUAN gắn đúng AAG/TC thay vì folder-token.
        # khoi: suy từ ĐƯỜNG DẪN nguồn (file_name có thể là path) -> stamp Khối cho dòng khoi NULL.
        # ĐỊNH DANH NGUỒN = '<công_ty_thư_mục>::<tên_file>' khi biết đường dẫn nguồn (dưới
        # received_reports) -> 3 nguồn trùng basename (vd B.6.XVP) KHÔNG đè nhau; else basename.
        _src = (SC.source_id_from_path(source_path) if source_path
                else os.path.basename(file_name or "")) or None   # tên file NGUỒN gốc
        _resolved = C.resolve_company(cong_ty, file_name)
        result["import"] = import_filled(out_path, cong_ty=_resolved,
                                         # khoi: tên file > path > KHỐI DUY NHẤT của pháp nhân (vd GA -> Công nghệ)
                                         # để dòng không bị khoi=NULL (mất khỏi filter Khối) — khoi_for_company
                                         # trả None nếu pháp nhân đa khối, nên không đóng dấu ẩu.
                                         khoi=khoi or C.khoi_from_path(source_path or file_name) or C.khoi_for_company(_resolved),
                                         source_file=_src)
    return result


def _source_sheets(data: bytes) -> list:
    wb = bb.fast_load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def guess_period(file_name: str):
    """Suy kỳ 'YYYY-MM' từ tên file: 'YYYYMM' (202601), 'MM.YYYY' (05.2026),
    hoặc 'THÁNG mm NĂM yyyy' (Báo cáo tiền tập đoàn). Đồng bộ sync_orchestrator._guess_period."""
    import re
    if not file_name:
        return None
    m = re.search(r"th[aá]ng\s*(\d{1,2})\s*n[aă]m\s*(20\d{2})", file_name, re.IGNORECASE)
    if m and 1 <= int(m.group(1)) <= 12:
        return f"{m.group(2)}-{int(m.group(1)):02d}"
    m = re.search(r"(20\d{2})[.\-_]?(0[1-9]|1[0-2])", file_name)   # 2026-01 / 202601
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"(0[1-9]|1[0-2])[.\-_](20\d{2})", file_name)     # 05.2026
    if m:
        return f"{m.group(2)}-{m.group(1)}"
    return None


def autofill_file(path: str, period: str = None, cong_ty: str = None, dry_run: bool = False) -> dict:
    """TỰ ĐỘNG (không LLM): với mỗi sheet nguồn có fill_spec ĐÃ HỌC (theo fingerprint) ->
    điền template + import raw_rows. Sheet chưa học -> bỏ qua (cần analyst). Dùng trong orchestrator.

    dry_run=True: CHỈ xem trước (row_count/sample/cảnh báo mỗi sheet), KHÔNG ghi file điền/KHÔNG
    nạp raw_rows — bước 0 rẻ (tất định, không LLM) để `analyst` biết ngay layout này ĐÃ học trước
    đó hay chưa, tránh phải phân tích lại từ đầu cho file/công ty đã quen (tăng tốc khi có nhiều
    file cùng layout xử lý liên tiếp)."""
    with open(path, "rb") as fh:
        data = fh.read()
    fname = os.path.basename(path)
    period = period or guess_period(fname)
    comp = C.resolve_company(cong_ty, fname)   # công ty của file (để tra spec scoped, tránh lây B11)
    processed, skipped = [], []
    headers = _all_sheet_headers(data)  # mở workbook 1 LẦN (file nặng: load rất chậm)
    for sheet, head in headers.items():
        if not head:
            continue
        cols = head[_header_row_of(head)]
        # Tra spec THEO CÔNG TY trước (spec có constants đặc thù công ty), rồi mới tới spec chung.
        spec = memory.fill_spec_find(memory.source_fingerprint(sheet, cols, scope=comp)) if comp else None
        if not spec:
            plain = memory.fill_spec_find(memory.source_fingerprint(sheet, cols))
            if plain and plain.get("constants"):
                # spec cũ (chung) nhưng mang constants đặc thù công ty -> KHÔNG áp mù cho công ty này;
                # cần analyst học lại theo công ty (fill_from_source sẽ lưu scoped).
                skipped.append(f"{sheet} (spec có constants công ty — học lại theo công ty)")
                continue
            spec = plain
        if spec and not C.role_allows(fname, spec.get("target_sheet")):
            # spec học được từ file vai khác (2 bản export cùng sổ) — không replay chéo vai.
            skipped.append(f"{sheet} (vai file '{C.file_role(fname)}' không nạp "
                           f"{spec.get('target_sheet')})")
            continue
        if not spec:
            skipped.append(sheet)
            continue
        r = fill_from_source(data, sheet, spec["target_sheet"], spec["mapping"],
                             period=period, cong_ty=cong_ty, file_name=fname,
                             dry_run=dry_run, auto_import=not dry_run, learn=False,
                             value_scale=spec.get("value_scale", 1.0),
                             constants=spec.get("constants"),
                             rename_rows=spec.get("rename_rows"),   # replay đổi tên dòng đã học -> KPI sáng
                             source_path=path)                      # để khối suy từ đường dẫn + source_file đúng
        entry = {"sheet": sheet, "target": spec["target_sheet"], "ok": r.get("ok")}
        if dry_run:
            entry.update(row_count=r.get("row_count"), sample=r.get("sample"),
                         unresolved_cc=r.get("unresolved_cc"), scale_warning=r.get("scale_warning"),
                         max_money_ty=r.get("max_money_ty"), source_fingerprint=r.get("source_fingerprint"))
        else:
            entry["rows"] = r.get("rows_written")
            entry["import"] = (r.get("import") or {}).get("ok")
        processed.append(entry)
    return {"ok": True, "file": fname, "period": period, "dry_run": dry_run,
            "processed": processed, "skipped_sheets": skipped,
            "any_processed": bool(processed)}

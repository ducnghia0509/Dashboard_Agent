# -*- coding: utf-8 -*-
"""Deriver CHUYÊN BIỆT: "TS hết khấu hao còn dùng / chờ thanh lý" (Khối 12, TSCĐ Chart 3/3).

Trước đây KHÔNG có nguồn tự động (xem memory tscd-chart3-hetkhauhao-blocked) -> payload
08_TSCD.het_kh_con_sd/het_kh_thanh_ly luôn 0. Nay có spec cột/sheet chính xác theo đơn vị
(user cung cấp 2026-07-29, đã đối chiếu khớp file thật kỳ T06/2026):

  · Layout PHẲNG (header 1 tầng): An Khách sạn, GA, Xanh VP, Trạm sạc, Dự án, SRVF (.xlsb) —
    lọc DÒNG TÀI SẢN có "Giá trị còn lại" (cuối kỳ) = 0, cộng "Nguyên giá" (cuối kỳ).
  · Layout NHÓM (header 2 tầng 'Tài sản'/'Khấu hao'/'Giá trị sổ sách'): An Taxi, Hưng Thịnh —
    dò cột CUỐI KỲ động (giống _derive_tscd._closing_col, KHÔNG cứng chữ cột vì lệch theo
    range ngày hiển thị mỗi tháng); dòng tài sản = có cột đặc tính (ngày mua lại/phương thức),
    dòng loại/tổng (vd "24213 Chi phí phân bổ CCDC") chỉ có cột A + số gộp -> bỏ.
  · HO: KHÔNG có sổ chi tiết tài sản để lọc GTCL=0 -> PROXY mã 223 BCĐKT (Giá trị hao mòn luỹ
    kế TOÀN BỘ, cuối kỳ) — khác bản chất chỉ tiêu (hao mòn lũy kế chung, không phải NG riêng
    của TS đã hết KH) nhưng là nguồn duy nhất khả dụng. Nguồn này là CÙNG FILE/CÙNG import_filled
    call với dòng "TSCĐ (theo CĐKT)" hiện có (_derive_tscd_cdkt trong agent_cli.py) -> ĐÃ gắn
    ngay trong đó (KHÔNG lặp lại ở đây, tránh 2 lần ghi cùng source_file đè lẫn nhau).
  · Chỉ An Khách sạn có cột "Tình trạng sử dụng" tách được còn dùng/chờ thanh lý (đã kiểm tra
    các đơn vị khác không có cột trạng thái tương đương) -> đơn vị khác đổ hết vào "còn sử dụng".
  · Depot PT/VP/TQ, HTX Xanh VP, HTX Xanh Tuyên Quang, XDV: CHƯA có nguồn -> bỏ qua (no-op).

API 2 lớp (tránh 2 lần import_filled CÙNG source_file đè nhau — xem agent_cli.py dispatcher):
  · compute(path) -> (con_sd, thanh_ly) RAW VND hoặc None — KHÔNG ghi DB.
  · extract(path, period, cong_ty) -> compute() rồi TỰ GHI (source_file riêng, 08_TSCD CHỈ 2
    cột hết-KH). Chỉ an toàn cho đơn vị mà _derive_tscd/_derive_tscd_duan KHÔNG ghi cho CHÍNH
    source_file này (An Khách sạn/GA/Xanh VP/Trạm sạc/Dự án/SRVF). An Taxi/Hưng Thịnh (layout
    NHÓM, _derive_tscd ghi thành công) -> dispatcher tự gắn kết quả compute() vào record của
    _derive_tscd, KHÔNG gọi extract()."""
import os
import sys

from openpyxl.utils import column_index_from_string as _ci

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers import template_filler as tf  # noqa: E402
from servers.common import be_bridge as bb  # noqa: E402
import agent_cli as A  # noqa: E402

norm = lambda v: bb.normalize_header(v, True) if v is not None else ""  # noqa: E731
_EPS = 1000.0   # 1.000đ — coi GTCL trong ngưỡng này là hết khấu hao (chống lệch làm tròn)


def _num(v):
    return v if isinstance(v, (int, float)) else None


def _is_zero(v):
    return v is not None and abs(v) < _EPS


# Layout PHẲNG — đối chiếu khớp file thật T06/2026 (xem docstring trên).
_FLAT = {
    "ANKHACHSAN": dict(sheet="danh sách sổ tài sản cố định", header_row=3, ng="J", gtcl="M", status="Q"),
    "GLOBALAI": dict(sheet="tài sản, ccdc", header_row=2, ng="D", gtcl="J"),
    "XANHVINHPHUC": dict(sheet="tài sản", header_row=6, ng="I", gtcl="AK"),
    "TRAMSAC": dict(sheet="biểu khấu hao", header_row=2, ng="I", gtcl="N"),
}
_DUAN = dict(header_row=6, ng="N", gtcl="P")
_GROUPED = {"ANTAXI", "HUNGTHINH"}
_SKIP_NAMES = ("tai san", "tong", "cong")


def _flat_rows(ws, cfg):
    hdr_r = cfg["header_row"]
    ng_i, gtcl_i = _ci(cfg["ng"]) - 1, _ci(cfg["gtcl"]) - 1
    status_i = _ci(cfg["status"]) - 1 if cfg.get("status") else None
    out = []
    for row in ws.iter_rows(min_row=hdr_r + 1, values_only=True):
        if not row or ng_i >= len(row) or gtcl_i >= len(row):
            continue
        ng, gtcl = _num(row[ng_i]), _num(row[gtcl_i])
        if ng is None or gtcl is None:
            continue
        if norm(row[0]) in _SKIP_NAMES:      # dòng tổng/tiêu đề (vd Trạm sạc dòng "Tài sản")
            continue
        status = row[status_i] if (status_i is not None and status_i < len(row)) else None
        out.append((ng, gtcl, status))
    return out


def _grouped_rows(rows, header_scan=8):
    grp_i = next((i for i, r in enumerate(rows[:header_scan])
                  if any("tai san" in norm(c) for c in r) and any("khau hao" in norm(c) for c in r)), None)
    if grp_i is None:
        return None
    group = A._forward_fill(rows[grp_i])
    sub = rows[grp_i + 1] if grp_i + 1 < len(rows) else []

    def _closing(kw):
        cands = [j for j in range(len(sub)) if j < len(group) and kw in norm(group[j])]
        if not cands:
            return None
        end = [j for j in cands if "31" in str(sub[j] or "") or "30" in str(sub[j] or "")]
        return (end or cands)[-1]
    ng_i = _closing("tai san")
    gtsl_i = _closing("gia tri so sach") or _closing("gia tri")
    if ng_i is None or gtsl_i is None:
        return None
    out = []
    for r in rows[grp_i + 2:]:
        if not r or all(c is None for c in r):
            continue
        # dòng CHI TIẾT tài sản có cột đặc tính (ngày mua lại/khấu hao lần đầu/phương thức, cột
        # B/C/D) — dòng loại/tổng (vd "24213 Chi phí phân bổ CCDC") chỉ có cột A + số gộp.
        has_detail = any(r[j] not in (None, "") for j in (1, 2, 3) if j < len(r))
        if not has_detail:
            continue
        ng = r[ng_i] if ng_i < len(r) else None
        gtsl = r[gtsl_i] if gtsl_i < len(r) else None
        ng, gtsl = _num(ng), _num(gtsl)
        if ng is None or gtsl is None:
            continue
        out.append((ng, gtsl, None))
    return out


def _split(rows_ngc):
    con_sd = thanh_ly = 0.0
    for ng, gtcl, status in rows_ngc:
        if not _is_zero(gtcl):
            continue
        st = norm(status)
        if "thanh ly" in st:
            thanh_ly += ng
        else:
            con_sd += ng
    return con_sd, thanh_ly


def _write(period, cong_ty, khoi, src, con_sd, thanh_ly):
    rec = [{"Kỳ": period, "Đơn vị": cong_ty,
            "TS hết KH còn sử dụng (NG, tỷ)": round(con_sd * 1e-9, 9),
            "TS hết KH chờ thanh lý (NG, tỷ)": round(thanh_ly * 1e-9, 9)}]
    out = os.path.join(tf.FILLED_DIR, f"TSHETKH_{period}_{cong_ty or 'NA'}_08_TSCD.xlsx")
    tf.fill("08_TSCD", rec, out)
    imp = tf.import_filled(out, cong_ty=cong_ty, khoi=khoi, source_file=src)
    return {"ok": bool(imp.get("rows_imported")), "rows": imp.get("rows_imported"),
            "con_sd": round(con_sd * 1e-9, 9), "thanh_ly": round(thanh_ly * 1e-9, 9)}


def compute(path):
    """Trả (con_sd, thanh_ly) RAW VND (chưa chia tỷ) từ báo cáo 'baocaotaisancodinhcongcudungcu',
    hoặc None nếu file/đơn vị này không khớp layout đã biết (HO: luôn None — xử lý riêng, gắn
    thẳng trong _derive_tscd_cdkt). Không ghi DB — dùng để GẮN vào record của _derive_tscd/
    _derive_tscd_duan (đơn vị mà 2 hàm đó đã ghi thành công cho CHÍNH source_file này, tránh 2 lần
    import_filled cùng source_file đè nhau) hoặc để extract() tự ghi standalone (đơn vị KHÔNG có
    dòng nào khác ghi cho source_file này)."""
    folder = A._source_id(path).split("::", 1)[0].upper()
    if folder == "HO":
        return None
    if folder == "SRVF" and path.lower().endswith(".xlsb"):
        try:
            return _compute_srvf_xlsb(path)
        except Exception:  # noqa: BLE001
            return None
    if folder not in _FLAT and folder != "DUAN" and folder not in _GROUPED:
        return None
    try:
        wb = bb.fast_load_workbook(path, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return None
    try:
        names = {s.strip().lower(): s for s in wb.sheetnames}
        rows_out = None
        if folder in _FLAT:
            cfg = _FLAT[folder]
            sn = names.get(cfg["sheet"])
            if sn:
                rows_out = _flat_rows(wb[sn], cfg)
        elif folder == "DUAN":
            sn = next((orig for low, orig in names.items()
                       if low.startswith("tháng") or low.startswith("thang")), None)
            if sn:
                rows_out = _flat_rows(wb[sn], _DUAN)
        elif folder in _GROUPED:
            sn = names.get("biểu khấu hao")
            if sn:
                rows = [list(r) for r in wb[sn].iter_rows(values_only=True)]
                rows_out = _grouped_rows(rows)
    finally:
        wb.close()
    if rows_out is None:
        return None
    return _split(rows_out)


def _compute_srvf_xlsb(path):
    from pyxlsb import open_workbook
    wb = open_workbook(path)
    try:
        sn = next((s for s in wb.sheets if norm(s).startswith("bao cao ts")), None)
        if not sn:
            return None
        with wb.get_sheet(sn) as sh:
            rows = []
            for row in sh.rows():
                cells = {c.c: c.v for c in row}
                maxc = max(cells) if cells else -1
                rows.append([cells.get(i) for i in range(maxc + 1)])
    finally:
        wb.close()
    hdr = next((i for i, r in enumerate(rows[:6])
                if any("loai tai san" in norm(c) for c in r)), None)
    if hdr is None:
        return None
    h = rows[hdr]

    def _col(name):
        return next((j for j, c in enumerate(h) if name in norm(c)), None)
    i_i, t_i, z_i = _col("loai tai san"), _col("nguyen gia cuoi ky"), _col("gia tri con lai")
    if None in (i_i, t_i, z_i):
        return None
    rows_out = []
    for r in rows[hdr + 1:]:
        if i_i >= len(r) or norm(r[i_i]) != "tscd":
            continue
        ng = r[t_i] if t_i < len(r) else None
        gtcl = r[z_i] if z_i < len(r) else None
        ng, gtcl = _num(ng), _num(gtcl)
        if ng is None or gtcl is None:
            continue
        rows_out.append((ng, gtcl, None))
    if not rows_out:
        return None
    return _split(rows_out)


def _khoi_of(path):
    """KHÔNG dùng A._khoi_of() thẳng: mọi file 'baocaotaisancodinhcongcudungcu' đặt tên
    'B.9.<mã>...' — số '9' trùng MÃ KHỐI 9 = "Khối hỗ trợ tập đoàn" trong quy ước
    khoi_from_filename (số 9 ở đây là SỐ THỨ TỰ BÁO CÁO trong bộ hồ sơ kế toán, không phải khối!)
    -> _khoi_of() ưu tiên filename sẽ gán NHẦM "Khối hỗ trợ tập đoàn" cho MỌI đơn vị (đã phát
    hiện 2026-07-29: An Taxi/An KS/GA/HT/XVP/Dự án/Trạm sạc/SRVF đều bị gộp nhầm vào khối HO khi
    lọc theo Khối trên FE). Dò theo ĐƯỜNG DẪN THƯ MỤC trước (đáng tin hơn cho báo cáo này); GA
    không map được qua path -> fallback bảng cứng theo khối THẬT (khớp dòng "TSCĐ theo CĐKT" của
    chính đơn vị đó, đã verify trong DB 2026-07-29)."""
    from servers.common import contract as C
    byp = C.khoi_from_path(path)
    if byp:
        return byp
    folder = A._source_id(path).split("::", 1)[0].upper()
    return {"GLOBALAI": "Khối KD Công nghệ"}.get(folder)


def extract(path, period, cong_ty=None):
    """Standalone: TỰ TÍNH + GHI (source_file riêng). Chỉ dùng khi CHẮC CHẮN không có deriver nào
    khác ghi report_type TS cho CHÍNH source_file này trong CÙNG lần chạy. HO (proxy gắn trong
    _derive_tscd_cdkt) không qua đây. Các đơn vị còn lại (kể cả An Taxi/Hưng Thịnh — dispatcher
    CHỦ ĐỘNG bỏ qua _derive_tscd cho các khối dùng CĐKT làm nguồn chính, tránh đếm đôi #17/#18
    Overview) đều ghi qua extract() này, source_file riêng, an toàn."""
    r = compute(path)
    if r is None:
        return {"ok": False, "skip": True}
    con_sd, thanh_ly = r
    khoi = _khoi_of(path)
    return _write(period, cong_ty, khoi, A._source_id(path), con_sd, thanh_ly)


if __name__ == "__main__":
    import argparse
    import json
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--cong-ty", dest="cong_ty", default=None)
    a = ap.parse_args()
    print(json.dumps(extract(a.file, a.period, a.cong_ty), ensure_ascii=False, default=str))

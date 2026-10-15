# -*- coding: utf-8 -*-
"""Đối chiếu "Giá trị tồn kho" màn vhkd4 (Showroom) với file nguồn KD23 (15/10/2026).

Câu hỏi thật đã làm bot bí: "dashboard 751 tỷ, file cũng là luỹ kế mà mới 400 tỷ". Hai con số đều đúng,
khác ĐỊNH NGHĨA: dòng "Tổng cộng" của file KD23 cộng cột "Dư cuối" của MỌI chân luân chuyển từ 01/01, kể
cả chân đã xuất (Tồn cuối = 0) mang dư ÂM (giá nhập ≠ giá xuất trên cùng số khung), còn dashboard chỉ cộng
xe ĐANG TỒN (Tồn cuối ≠ 0). Tool tách đúng các phần đó ra từ chính file, đặt cạnh số DB và số dashboard.
Chỉ đọc."""
import glob
import os

from . import so_lieu_check as sc

THU_MUC = os.path.join(sc.RECEIVED_DIR, "TEST_SR", "baocaonhapxuatxetheosokhung")
_TY = 1e9


def _file_ngay(ngay):
    """File KD23 của đúng ngày; không có thì file gần nhất TRƯỚC ngày đó (kèm cờ)."""
    ds = sorted(glob.glob(os.path.join(THU_MUC, "*_Baocaonhapxuatxetheosokhung.xlsx")))
    theo_ngay = {}
    for p in ds:
        n = sc._ngay_tu_ten(os.path.basename(p).replace("_", "."))
        if n:
            theo_ngay[n] = p
    if ngay in theo_ngay:
        return theo_ngay[ngay], ngay
    truoc = [n for n in theo_ngay if n <= ngay]
    return (theo_ngay[max(truoc)], max(truoc)) if truoc else (None, None)


def _doc_file(p):
    import openpyxl
    ws = openpyxl.load_workbook(p, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    hdr = None
    for i, r in enumerate(rows, 1):
        if r and any(str(v or "").strip() == "Tồn cuối" for v in r):
            hdr = {str(v).strip(): j for j, v in enumerate(r) if v}
            break
        if i > 30:
            break
    if not hdr:
        raise ValueError("không thấy dòng tiêu đề có cột 'Tồn cuối'")
    K, O, C = hdr["Tồn cuối"], hdr["Dư cuối"], hdr["Số khung"]
    tong = None
    nhom = {}          # tồn cuối -> [số dòng, Σ dư]
    vin_nhieu = {}
    for r in rows:
        r = tuple(r) + (None,) * 40
        if not r[C]:
            if tong is None and isinstance(r[K], (int, float)):
                tong = {"ton_cuoi": r[K], "du_cuoi_ty": (r[O] or 0) / _TY}   # dòng "Tổng cộng" ngay dưới header
            continue
        k = r[K] or 0
        g = nhom.setdefault(k, [0, 0.0])
        g[0] += 1
        g[1] += r[O] or 0
        if k:
            vin_nhieu.setdefault(r[C], []).append(k)
    return tong, nhom, {v: x for v, x in vin_nhieu.items() if len(x) > 1}


def ton_kho_xe(ngay: str, env: str = "prod") -> dict:
    p, ngay_file = _file_ngay(ngay)
    if not p:
        return {"loi": f"không có file KD23 nào tới ngày {ngay} trong {THU_MUC}"}
    tong, nhom, vin_nhieu = _doc_file(p)
    xe_ton = sum(k * v[0] for k, v in nhom.items())          # RÒNG: chân −1 trừ đi (= Σ cột Tồn cuối)
    du_ton = sum(v[1] for k, v in nhom.items() if k != 0) / _TY
    du_da_xuat = nhom.get(0, [0, 0.0])[1] / _TY
    db = sc._ro(env, "SELECT count(*) AS dong, sum(amount) AS sl, sum(amount2) AS gt, count(distinct payload::json->>'so_khung') AS vin "
                     "FROM raw_rows WHERE report_type='VHKD_TONHD' AND ngay=%s", (ngay_file,))
    db = db[0] if db else {}
    try:
        from . import man_hien_thi as mh
        d = (mh.lay_man("vhkd4", ngay_file[:7], env, "day", ngay_file[:8] + "01", ngay_file) or {}).get("vhkd4") or {}
        dash = {"xe": d.get("tonHoaDonXe"), "gia_tri_ty": d.get("tonHoaDonGiaTri"), "ngay_chot": d.get("ngayChotTheoKenh")}
    except Exception as e:  # noqa: BLE001
        dash = {"loi": str(e)[:200]}
    out = {
        "ngay_hoi": ngay, "ngay_file": ngay_file, "file": os.path.basename(p),
        "file_dong_tong_cong": {"xe": tong and tong["ton_cuoi"], "du_cuoi_ty": tong and round(tong["du_cuoi_ty"], 4)},
        "file_xe_dang_ton": {"xe": xe_ton, "du_cuoi_ty": round(du_ton, 4)},
        "file_chan_da_xuat_ton_cuoi_0": {"dong": nhom.get(0, [0])[0], "du_cuoi_ty": round(du_da_xuat, 4)},
        "file_chan_ton_cuoi_am": {"dong": sum(v[0] for k, v in nhom.items() if k < 0),
                                  "so_khung_nhieu_chan": len(vin_nhieu), "vi_du": list(vin_nhieu.items())[:3]},
        "db": {"dong": db.get("dong"), "sl": sc._f(db.get("sl")), "gia_tri_ty": sc._f(db.get("gt")), "so_khung": db.get("vin")},
        "dashboard_vhkd4": dash,
        "chenh_dashboard_vs_file_xe_dang_ton_ty": (round(dash["gia_tri_ty"] - du_ton, 4)
                                                   if isinstance(dash.get("gia_tri_ty"), (int, float)) else None),
        "chenh_xe_dashboard_vs_file": (dash["xe"] - xe_ton if isinstance(dash.get("xe"), int) else None),
        "chenh_dashboard_vs_dong_tong_cong_ty": (round(dash["gia_tri_ty"] - tong["du_cuoi_ty"], 4)
                                                 if tong and isinstance(dash.get("gia_tri_ty"), (int, float)) else None),
        "giai_thich": (
            "Giá trị tồn kho trên dashboard = Σ 'Dư cuối' của XE ĐANG TỒN (Tồn cuối ≠ 0) tại NGÀY CHỐT — là ẢNH CHỤP, "
            "không cộng dồn các ngày. Dòng 'Tổng cộng' của file cộng cả chân ĐÃ XUẤT (Tồn cuối = 0) vốn mang dư ÂM, "
            "nên Tổng cộng = xe đang tồn + chân đã xuất (số âm) và luôn thấp hơn dashboard. Người dùng so với dòng "
            "Tổng cộng là so hai định nghĩa khác nhau. Khớp đúng là: dashboard ≈ file_xe_dang_ton."),
    }
    if ngay_file != ngay:
        out["canh_bao"] = f"không có file ngày {ngay}, đang dùng file gần nhất {ngay_file}"
    return out

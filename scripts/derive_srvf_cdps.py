# -*- coding: utf-8 -*-
"""Deriver CHUYÊN BIỆT cho SRVF (CHI NHÁNH VINFAST) — CĐPS 1 TẦNG (format riêng, deriver chung
không đọc được). Sheet 'CĐPS': A=Tài khoản · B=Tên · C=Nợ đầu · D=Có đầu · E=PS Nợ · F=PS Có ·
G=Dư nợ cuối · H=Dư có cuối. Bóc theo TK (chị Điệp: công nợ CĐPS TK131/331):
  · TK 131 -> PTHU (05): dư nợ (đầu C / cuối G, tăng=PS Nợ E, giảm=PS Có F).
  · TK 331 -> PTRA (06): dư có (đầu D / cuối H, tăng=PS Có F, giảm=PS Nợ E).

cong_ty=TC, khối Vinfast Showroom (SRVF thuộc Cổ phần Thịnh Cường). Idempotent theo source_file.
Chạy: .venv/bin/python scripts/derive_srvf_cdps.py <file.xlsx> --period 2026-06
"""
import argparse
import os
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers import template_filler as tf  # noqa: E402
from servers.common import be_bridge as bb  # noqa: E402
import agent_cli as A  # noqa: E402

norm = lambda v: bb.normalize_header(v, True)  # noqa: E731
TY = lambda v: round(v * 1e-9, 9) if isinstance(v, (int, float)) else None  # noqa: E731


def _cols(rows):
    """Dò dòng HEADER CỘT (không phải dòng tiêu đề) + index cột theo VAI. Header cột = dòng mà cột
    'phát sinh nợ' + 'dư nợ cuối' dò được thành ô RIÊNG (dòng tiêu đề gộp 1 ô -> loại)."""
    for i, r in enumerate(rows[:12]):
        low = [norm(c) for c in r]
        def col(*kw):
            return next((j for j, c in enumerate(low) if all(k in c for k in kw)), None)
        m = {"tk": col("tai khoan"), "no_dau": col("no dau"), "co_dau": col("co dau"),
             "ps_no": col("phat sinh no"), "ps_co": col("phat sinh co"),
             "no_cuoi": col("du no cuoi"), "co_cuoi": col("du co cuoi")}
        if m["ps_no"] is not None and m["no_cuoi"] is not None:   # đúng dòng header cột
            return i, m
    return None, None


def _cdkt_ma(wbk, norm, ma):
    """Đọc 1 mã số CĐKT (Số đầu kỳ, Số cuối kỳ) từ sheet CĐKT chuẩn TT200 (cột 'Số cuối kỳ'/'Số đầu
    kỳ'). Trả (dau, cuoi) VND hoặc (None, None) nếu không thấy sheet/mã."""
    sn = next((s for s in wbk.sheetnames if "cdkt" in norm(s) or "can doi ke toan" in norm(s)
               or "đkt" in s.lower()), None)
    if sn is None:
        return (None, None)
    rows = [list(r) for r in wbk[sn].iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:15])
               if any(norm(c) == "ma so" for c in r) and any("cuoi" in norm(c) for c in r)), None)
    if hi is None:
        return (None, None)
    hdr = rows[hi]
    ma_j = next((j for j, c in enumerate(hdr) if norm(c) == "ma so"), None)
    cuoi_j = next((j for j, c in enumerate(hdr) if "cuoi" in norm(c)), None)
    dau_j = next((j for j, c in enumerate(hdr) if "dau" in norm(c) and "nam" not in norm(c)), None)
    if ma_j is None or cuoi_j is None:
        return (None, None)
    for r in rows[hi + 1:]:
        if ma_j < len(r) and str(r[ma_j]).strip() == ma:
            _num = lambda j: r[j] if (j is not None and j < len(r) and isinstance(r[j], (int, float))) else None
            return (_num(dau_j), _num(cuoi_j))
    return (None, None)


def _cdkt_ma140(wbk, norm):
    """Mã 140 'Hàng tồn kho' (đầu, cuối) — dùng làm TỔNG tồn kho THẬT để cân dòng '156_xe' thiếu."""
    return _cdkt_ma(wbk, norm, "140")


def extract(path, period, cong_ty="TC"):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "CĐPS" not in wb.sheetnames:
        return {"ok": False, "error": "không thấy sheet CĐPS"}
    rows = [list(r) for r in wb["CĐPS"].iter_rows(values_only=True)]
    # SỐ DƯ công nợ lấy từ CĐKT (hướng dẫn C Điệp cập nhật: #30 phải thu = CĐKT Mã 131; #36 phải trả
    # = CĐKT Mã 311). CĐPS chỉ dùng cho PS tăng/giảm (#31/#37). CĐPS≠CĐKT ở tháng có phân loại lệch
    # (vd T05: CĐPS 131=198,464 vs CĐKT=198,460). None -> fallback CĐPS (giữ tương thích cũ).
    _ck131_dau, _ck131_cuoi = _cdkt_ma(wb, norm, "131")   # phải thu KH ngắn hạn
    _ck311_dau, _ck311_cuoi = _cdkt_ma(wb, norm, "311")   # phải trả người bán ngắn hạn
    wb.close()
    hi, c = _cols(rows)
    if hi is None:
        return {"ok": False, "error": "không dò được header CĐPS"}
    src, khoi = A._source_id(path), A._khoi_of(path)

    def find_tk(tk):
        for r in rows[hi + 1:]:
            if r and str(bb.parse_text(r[c["tk"]])).strip() == tk:
                return r
        return None

    def val(r, key):
        j = c[key]
        return TY(r[j]) if (j is not None and j < len(r)) else None

    out = {"ok": True}
    # PTHU (TK 131, dư NỢ)
    r131 = find_tk("131")
    if r131:
        rec = [{"Kỳ": period, "Đơn vị": cong_ty, "Khách hàng": "Phải thu khách hàng (tổng)",
                "Dư cuối kỳ (tỷ)": TY(_ck131_cuoi) if _ck131_cuoi is not None else val(r131, "no_cuoi"),
                "Dư đầu kỳ (tỷ)": TY(_ck131_dau) if _ck131_dau is not None else val(r131, "no_dau"),
                "PS tăng - Nợ (tỷ)": val(r131, "ps_no"), "PS giảm - Có (tỷ)": val(r131, "ps_co")}]
        p = os.path.join(tf.FILLED_DIR, f"SRVF_{period}_05_PHAITHU.xlsx")
        tf.fill("05_PHAITHU", rec, p)
        out["pthu"] = tf.import_filled(p, cong_ty=cong_ty, khoi=khoi, source_file=src).get("rows_imported")
    # PTRA (TK 331, dư CÓ)
    r331 = find_tk("331")
    if r331:
        rec = [{"Kỳ": period, "Đơn vị": cong_ty, "Nhà cung cấp": "Phải trả nhà cung cấp (tổng)",
                "Dư cuối kỳ (tỷ)": TY(_ck311_cuoi) if _ck311_cuoi is not None else val(r331, "co_cuoi"),
                "Dư đầu kỳ (tỷ)": TY(_ck311_dau) if _ck311_dau is not None else val(r331, "co_dau"),
                "PS tăng - Có (tỷ)": val(r331, "ps_co"), "PS giảm - Nợ (tỷ)": val(r331, "ps_no")}]
        p = os.path.join(tf.FILLED_DIR, f"SRVF_{period}_06_PHAITRA.xlsx")
        tf.fill("06_PHAITRA", rec, p)
        out["ptra"] = tf.import_filled(p, cong_ty=cong_ty, khoi=khoi, source_file=src).get("rows_imported")
    # CHIỀU ĐẢO công nợ (ADDITIVE, cho thẻ "Trả trước NCC"/"Người mua trả tiền trước" màn Công nợ —
    # FE đọc report_type PTRA_ADV/PTHU_ADV): TK331 dư NỢ cuối = mình ứng trước cho NCC (mã 132);
    # TK131 dư CÓ cuối = khách ứng trước cho mình (mã 312/321). Nguồn CĐPS (khớp CĐKT). Insert trực
    # tiếp (report_type ADDITIVE, không có template); idempotent theo source_file+period+report_type.
    _adv = [("PTRA_ADV", r331, "no_cuoi", "Trả trước NCC (tổng)"),         # TK331 dư Nợ = ứng trước NCC
            ("PTHU_ADV", r131, "co_cuoi", "Người mua trả tiền trước (tổng)")]  # TK131 dư Có = khách ứng trước
    _dbh = bb.db.get_db()
    for _rt, _r, _key, _lbl in _adv:
        if not _r:
            continue
        _v = val(_r, _key)
        _dbh.execute("DELETE FROM raw_rows WHERE source_file=? AND period_month=? AND report_type=?",
                     (src, period, _rt))
        if _v is None or abs(_v) < 1e-9:
            continue
        _twin = _dbh.execute(
            "SELECT dataset_id, ngay FROM raw_rows WHERE source_file=? AND period_month=? "
            "AND report_type IN ('PTRA','PTHU') LIMIT 1", (src, period)).fetchone()
        if not _twin:
            continue
        import json as _json
        _dbh.execute(
            "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,cost_center,"
            "period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (_twin["dataset_id"], _rt, 6000000, _twin["ngay"], cong_ty, khoi, None, period,
             _v, None, _lbl, None, None,
             _json.dumps({"unit": "ty", "nguon": f"CĐPS {_key}"}, ensure_ascii=False), src))
        out[_rt.lower()] = _v
    _dbh.commit()
    # THUẾ: TK 133 (GTGT được khấu trừ = phải thu, dư NỢ cuối) + TK 333 (thuế phải nộp, dư CÓ cuối)
    thue = []
    for tk, pt, key in (("133", "Phải thu", "no_cuoi"), ("333", "Phải nộp", "co_cuoi")):
        r = find_tk(tk)
        if r and val(r, key):
            ten = bb.parse_text(r[c["tk"] + 1]) if c["tk"] + 1 < len(r) else None
            thue.append({"Kỳ": period, "Đơn vị": cong_ty,
                         "Loại thuế (GTGT ra/vào, TNCN, TNDN, NK, khác)": ten or f"TK {tk}",
                         "Phải thu/Phải nộp": pt, "Dư cuối kỳ (tỷ)": val(r, key)})
    if thue:
        p = os.path.join(tf.FILLED_DIR, f"SRVF_{period}_10_THUE.xlsx")
        tf.fill("10_THUE", thue, p)
        out["thue"] = tf.import_filled(p, cong_ty=cong_ty, khoi=khoi, source_file=src).get("rows_imported")

    # TỒN KHO xe VinFast: sheet 156_xe, gộp theo MODEL (cột 'Mã kx'). Cột: Dư đầu=L, Giá trị nhập=M,
    # Giá trị xuất=N, Tồn cuối=O (=L+M-N). Giá trị theo VND.
    wb2 = openpyxl.load_workbook(path, data_only=True, read_only=True)
    # Sheet tồn kho xe VinFast ĐỔI TÊN theo tháng: '156_xe' (T03+), '156 xe' (T01), '156' (T02).
    # Trước khớp CỨNG '156_xe' -> T01/T02 MẤT tồn kho. Nhận theo TÊN chứa '156' (loại '156_khác') +
    # CHỮ KÝ header ('ma kx'+'ton cuoi') để chắc đúng sheet xe (không nhầm '156_khác'/'152'…).
    _inv_sheet, xr = None, None
    for _s in wb2.sheetnames:
        _ns = norm(_s)
        if "156" not in _ns or "khac" in _ns:
            continue
        _xr = [list(r) for r in wb2[_s].iter_rows(values_only=True)]
        if next((i for i, r in enumerate(_xr[:12])
                 if any("ma kx" in norm(x) for x in r) and any("ton cuoi" in norm(x) for x in r)), None) is not None:
            _inv_sheet, xr = _s, _xr
            break
    if _inv_sheet is not None:
        # header: dò dòng có 'ma kx' + 'ton cuoi'
        xhi = next((i for i, r in enumerate(xr[:12])
                    if any("ma kx" in norm(x) for x in r) and any("ton cuoi" in norm(x) for x in r)), None)
        if xhi is not None:
            low = [norm(x) for x in xr[xhi]]
            cx = {"model": next((j for j, x in enumerate(low) if "ma kx" in x), None),
                  "dau": next((j for j, x in enumerate(low) if x == "du dau" or "du dau" in x), None),
                  "nhap": next((j for j, x in enumerate(low) if "gia tri nhap" in x), None),
                  "xuat": next((j for j, x in enumerate(low) if "gia tri xuat" in x), None)}
            agg = {}
            for r in xr[xhi + 1:]:
                m = bb.parse_text(r[cx["model"]]) if (cx["model"] is not None and cx["model"] < len(r)) else None
                if not m or norm(m).startswith(("tong", "cong")):
                    continue
                a = agg.setdefault(str(m).strip(), {"dau": 0.0, "nhap": 0.0, "xuat": 0.0})
                for k in ("dau", "nhap", "xuat"):
                    v = r[cx[k]] if (cx[k] is not None and cx[k] < len(r) and isinstance(r[cx[k]], (int, float))) else 0
                    a[k] += v
            recs = []
            sum_dau = sum_cuoi = 0.0
            for m, a in agg.items():
                cuoi = a["dau"] + a["nhap"] - a["xuat"]
                sum_dau += a["dau"]
                sum_cuoi += cuoi
                if abs(cuoi) < 1e3 and not a["nhap"] and not a["xuat"]:
                    continue
                recs.append({"Kỳ": period, "Đơn vị": cong_ty, "Loại HTK (NVL/Vật tư/Hàng hóa…)": m,
                             "TK (151-156)": "156", "Dư đầu kỳ (tỷ)": TY(a["dau"]),
                             "Nhập trong kỳ (tỷ)": TY(a["nhap"]), "Xuất trong kỳ (tỷ)": TY(a["xuat"]),
                             "Dư cuối kỳ (tỷ)": TY(cuoi)})
            # DÒNG CÂN BẰNG "khác": 156_xe CHỈ là xe (⊂ TK156). Tổng tồn kho THẬT = CĐKT mã140 =
            # TK152(NVL)+153(CCDC)+154(SPDD)+156(xe+phụ kiện). Thêm 1 dòng = mã140 − Σ156_xe để tổng
            # màn Tồn kho KHỚP mã140 (trước chỉ 156_xe -> thiếu ~90 tỷ). Đọc mã140 đầu/cuối từ CĐKT.
            _m140_dau, _m140_cuoi = _cdkt_ma140(wb2, norm)
            if _m140_dau is not None or _m140_cuoi is not None:
                _khac_dau = (_m140_dau - sum_dau) if _m140_dau is not None else 0.0
                _khac_cuoi = (_m140_cuoi - sum_cuoi) if _m140_cuoi is not None else 0.0
                if _khac_dau > 1e6 or _khac_cuoi > 1e6:   # có tồn kho ngoài xe (NVL/CCDC/SPDD/phụ kiện)
                    recs.append({"Kỳ": period, "Đơn vị": cong_ty,
                                 "Loại HTK (NVL/Vật tư/Hàng hóa…)": "Vật tư, CCDC, SPDD & phụ kiện (ngoài xe)",
                                 "TK (151-156)": "152-156", "Dư đầu kỳ (tỷ)": TY(_khac_dau),
                                 "Dư cuối kỳ (tỷ)": TY(_khac_cuoi)})
            if recs:
                p = os.path.join(tf.FILLED_DIR, f"SRVF_{period}_09_TONKHO.xlsx")
                tf.fill("09_TONKHO", recs, p)
                out["tonkho"] = tf.import_filled(p, cong_ty=cong_ty, khoi=khoi, source_file=src).get("rows_imported")
    wb2.close()

    if not any(k in out for k in ("pthu", "ptra", "thue", "tonkho")):
        return {"ok": False, "error": "không bóc được gì"}
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    a = ap.parse_args()
    import json
    print(json.dumps(extract(a.file, a.period), ensure_ascii=False))

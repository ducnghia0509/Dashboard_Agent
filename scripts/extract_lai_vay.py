# -*- coding: utf-8 -*-
"""Bù "LÃI VAY TRONG KỲ" cho màn Dòng tiền/Vay — ADDITIVE, giữ nguyên cách cũ.

Nguồn: sổ DATA (định khoản) trong báo cáo ngân hàng đơn vị (THUCHI/baocaonganhang) —
bút toán CHI có 'lãi' trong diễn giải, lọc theo THÁNG (DATA lũy kế đầu năm), gộp theo
Mã ngân hàng. (Cột 'Lãi vay đã trả' trong sheet NH/TH là LŨY KẾ nên KHÔNG dùng.)

GHI: MERGE patch `amount2` (lãi vay trong kỳ, TỶ) vào ĐÚNG dòng report_type='VAY' hiện có
theo (period_month, cong_ty, dim1=ngân hàng). Không tạo dòng trùng, không đụng dư nợ.
metrics_extra.vay_extras đọc amount2 -> lai / tlLaiDt / laiByBank + card "Lãi vay".

Chỉ TC/XVP/VFQN/AN (có DATA thật). HungThinh bỏ (DATA là bản sao Thịnh Cường).

Chạy:  .venv/bin/python scripts/extract_lai_vay.py            # dry-run
       .venv/bin/python scripts/extract_lai_vay.py --commit
"""
import argparse
import datetime
import glob
import json
import os
import sys

sys.path.insert(0, "/home/sysadmin")
import cashflow_vay_extractor as ext            # noqa: E402  (unit_of/month_of/load)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers.common import be_bridge as bb      # noqa: E402

MASTER = {"ThinhCuong": "TC", "XanhVP": "XVP", "QuangNinh": "VFQN", "AN": "AAG"}
SKIP_UNITS = {"HungThinh", "HTX_Xanh"}           # DATA không đáng tin / không có


def bank_of(ma):
    """Mã ngân hàng trong DATA -> dim1 của dòng VAY."""
    s = str(ma or "").strip().upper()
    if s in ("B1", "B2", "B3", "B2B"): return "BIDV"   # TC: tài khoản BIDV Sơn Tây
    if "BAB" in s or "BAC A" in s or "BẮC" in s: return "Bắc Á"
    if "BIDV" in s: return "BIDV"
    if "VCB" in s: return "VCB"
    if "VPB" in s or "VP" in s: return "VP"
    if "MB" in s: return "MB"
    return None


def _cidx(H, *kw):
    return next((j for j, c in enumerate(H) if isinstance(c, str) and all(k in c.lower() for k in kw)), None)


def lai_by_bank(wb, month, year=2026):
    """{bank_dim1: lãi_trong_kỳ(VND)} từ DATA, lọc tháng, CHI có 'lãi'."""
    if "DATA" not in wb.sheetnames:
        return {}, 0.0
    ws = wb["DATA"]
    rs = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
    if len(rs) < 2:
        return {}, 0.0
    H = rs[1]
    c_dg, c_chi, c_date, c_nh = _cidx(H, "diễn giải"), _cidx(H, "chi"), _cidx(H, "ngày"), _cidx(H, "ngân hàng")
    if None in (c_dg, c_chi, c_date, c_nh):
        return {}, 0.0
    out, unmapped = {}, 0.0
    for r in ws.iter_rows(min_row=3, values_only=True):
        d = r[c_date] if c_date < len(r) else None
        if not (isinstance(d, datetime.datetime) and d.year == year and d.month == month):
            continue
        dg = str(r[c_dg]).lower() if c_dg < len(r) and r[c_dg] else ""
        chi = r[c_chi] if c_chi < len(r) and isinstance(r[c_chi], (int, float)) else 0
        if chi and "lãi" in dg:
            b = bank_of(r[c_nh] if c_nh < len(r) else None)
            if b:
                out[b] = out.get(b, 0.0) + chi
            else:
                unmapped += chi
    return out, unmapped


def compute():
    """{(cong_ty, period, bank): lãi(VND)}"""
    plan, warns = {}, []
    for f in sorted(glob.glob(os.path.join(ext.DATA_DIR, "*.xlsx"))):
        unit = ext.unit_of(os.path.basename(f))
        month = ext.month_of(os.path.basename(f))
        if unit in SKIP_UNITS or unit not in MASTER or not month:
            continue
        wb = ext.load(f)
        lai, unmapped = lai_by_bank(wb, month)
        wb.close()
        for b, v in lai.items():
            plan[(MASTER[unit], f"2026-{month:02d}", b)] = plan.get((MASTER[unit], f"2026-{month:02d}", b), 0.0) + v
        if unmapped > 1e8:
            warns.append(f"{unit} 2026-{month:02d}: lãi chưa map ngân hàng {unmapped/1e9:.2f} tỷ")
    return plan, warns


def merge(plan, commit):
    db = bb.db.get_db()
    matched = unmatched = 0
    for (cty, period, bank), val in sorted(plan.items()):
        rows = db.execute(
            "SELECT id, amount2 FROM raw_rows WHERE report_type='VAY' "
            "AND period_month=? AND cong_ty=? AND dim1=?", (period, cty, bank)).fetchall()
        if not rows:
            unmatched += 1
            print(f"  [KHÔNG KHỚP] {period} {cty}/{bank}: lãi={val/1e9:.3f} tỷ (chưa có dòng VAY)")
            continue
        val_ty = val / 1e9
        for r in rows:
            d = dict(r) if not isinstance(r, dict) else r
            if commit:
                db.execute("UPDATE raw_rows SET amount2=? WHERE id=?", (val_ty, d["id"]))
            print(f"  {'[GHI]' if commit else '[DỰ KIẾN]'} {period} {cty}/{bank}: "
                  f"lãi {(d.get('amount2') or 0):.3f} -> {val_ty:.3f} tỷ")
            matched += 1
    if commit:
        db.commit()
    return matched, unmatched


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    a = ap.parse_args()
    plan, warns = compute()
    print(f"== Lãi vay trong kỳ: {len(plan)} khoá (cong_ty,period,bank) ==")
    matched, unmatched = merge(plan, a.commit)
    for w in warns:
        print("  ⚠️", w)
    print(f"\n{'ĐÃ GHI' if a.commit else 'DRY-RUN'}: khớp {matched}, không khớp {unmatched}. "
          f"{'' if a.commit else 'Chạy lại với --commit để ghi.'}")


if __name__ == "__main__":
    main()

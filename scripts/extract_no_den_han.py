# -*- coding: utf-8 -*-
"""Dẫn xuất "NỢ GỐC ĐẾN HẠN TRONG THÁNG" cho màn Dòng tiền (phần Vay) — ADDITIVE.

Nguồn: báo cáo ngân hàng đơn vị trong THUCHI/baocaonganhang (spec:
/home/sysadmin/nguyen_tac_lay_no_den_han_dashboard.md). Tính = ma trận cột-tháng của
sheet NH/TH (đọc cột = tháng của file) + Thịnh Cường cột rộng 'Trả gốc T<MM>'.

GHI: MERGE (không tạo dòng trùng) — patch payload.den_han vào ĐÚNG dòng report_type='VAY'
đang có (do extract_sodutien_vay tạo) theo khoá (period_month, cong_ty, dim1=ngân hàng).
Chỉ đụng trường den_han; KHÔNG động dư nợ đầu/vay/trả/cuối. metrics_extra.vay_extras SUM
payload.den_han -> denHan/tlDenHan/bankDenHan.

Chạy:
  .venv/bin/python scripts/extract_no_den_han.py            # DRY-RUN (chỉ in kế hoạch)
  .venv/bin/python scripts/extract_no_den_han.py --commit   # ghi DB (patch den_han)
"""
import argparse
import glob
import json
import os
import sys

sys.path.insert(0, "/home/sysadmin")               # extractor lõi (đã verify)
import cashflow_vay_extractor as ext                # noqa: E402
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers.common import be_bridge as bb          # noqa: E402

# đơn vị (tên file) -> mã công ty master
MASTER = {"ThinhCuong": "TC", "HungThinh": "HT", "XanhVP": "XVP",
          "QuangNinh": "VFQN", "AN": "AAG", "HTX_Xanh": "HTX_XVP"}


def bank_of(sheet_name):
    """Sheet NH/TH -> tên ngân hàng khớp dim1 của dòng VAY hiện có."""
    s = sheet_name.upper()
    if "BAB" in s:                                   return "Bắc Á"
    if "BIDV" in s:                                  return "BIDV"
    if "-MB" in s or s.endswith(" MB") or "MB " in s: return "MB"
    if "VPB" in s or "VFQN" in s:                    return "VP"
    if "TRUNG HẠN 2022" in s:                        return "BIDV"   # TC: gộp về BIDV
    return None


def compute_plan(only_period=None):
    """{(cong_ty, period, bank): den_han(VND)} từ báo cáo ngân hàng. only_period='2026-MM' -> chỉ kỳ đó
    (autofill gọi scope 1 kỳ để KHÔNG đụng kỳ khác)."""
    plan, unmapped = {}, []
    for f in sorted(glob.glob(os.path.join(ext.DATA_DIR, "*.xlsx"))):
        unit = ext.unit_of(os.path.basename(f))
        month = ext.month_of(os.path.basename(f))
        master = MASTER.get(unit)
        if not (unit and month and master):
            continue
        period = f"2026-{month:02d}"
        if only_period and period != only_period:
            continue
        wb = ext.load(f)
        for nm in wb.sheetnames:
            if nm.startswith("foxz"):
                continue
            if ext.sheet_kind(wb[nm]) in ("NH", "TH"):
                v = ext.matrix_month(wb[nm], month)
                if v:
                    b = bank_of(nm)
                    if not b:
                        unmapped.append((unit, nm)); continue
                    plan[(master, period, b)] = plan.get((master, period, b), 0.0) + v
        if unit == "ThinhCuong":
            for nm in ("Trung hạn 2022", "Trung hạn BIDV"):
                if nm in wb.sheetnames:
                    v, _ = ext.wide_tra_goc(wb[nm], month)
                    if v:
                        plan[(master, period, "BIDV")] = plan.get((master, period, "BIDV"), 0.0) + v
        wb.close()
    return plan, unmapped


def merge(plan, commit):
    db = bb.db.get_db()
    matched = unmatched = 0
    for (cty, period, bank), val in sorted(plan.items()):
        rows = db.execute(
            "SELECT id, payload FROM raw_rows WHERE report_type='VAY' "
            "AND period_month=? AND cong_ty=? AND dim1=?",
            (period, cty, bank)).fetchall()
        if not rows:
            unmatched += 1
            print(f"  [KHÔNG KHỚP] {period} {cty}/{bank}: den_han={val/1e9:.3f} tỷ "
                  f"(chưa có dòng VAY tương ứng)")
            continue
        val_ty = val / 1e9   # nguồn ở ĐỒNG; template/metrics VAY dùng TỶ (khớp amount/du_dau_ky…)
        for r in rows:
            d = dict(r) if not isinstance(r, dict) else r
            pl = json.loads(d.get("payload") or "{}")
            old = pl.get("den_han")
            pl["den_han"] = val_ty
            if commit:
                db.execute("UPDATE raw_rows SET payload=? WHERE id=?",
                           (json.dumps(pl, ensure_ascii=False), d["id"]))
            print(f"  {'[GHI]' if commit else '[DỰ KIẾN]'} {period} {cty}/{bank}: "
                  f"den_han {(old or 0):.3f} -> {val_ty:.3f} tỷ")
            matched += 1
    if commit:
        db.commit()
    return matched, unmatched


def apply(only_period=None, commit=True):
    """Autofill hook: patch den_han cho 1 kỳ. Trả (matched, unmatched)."""
    plan, _ = compute_plan(only_period)
    return merge(plan, commit)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    plan, unmapped = compute_plan()
    print(f"== Kế hoạch nợ đến hạn: {len(plan)} khoá (cong_ty,period,bank) ==")
    matched, unmatched = merge(plan, a.commit)
    if unmapped:
        print(f"\n⚠️ Sheet không map được ngân hàng (bỏ qua): {sorted(set(unmapped))}")
    print(f"\n{'ĐÃ GHI' if a.commit else 'DRY-RUN'}: khớp {matched}, không khớp {unmatched}. "
          f"{'' if a.commit else 'Chạy lại với --commit để ghi.'}")


if __name__ == "__main__":
    main()

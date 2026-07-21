# -*- coding: utf-8 -*-
"""Tách dòng VAY theo KỲ HẠN (Ngắn/Trung hạn) — REBUILD, giữ tổng khớp group report.

Bối cảnh: group report (Báo cáo tiền tập đoàn) cấp dư nợ đầu/vay/trả/cuối per (đơn vị,
ngân hàng) NHƯNG KHÔNG có kỳ hạn. Báo cáo ngân hàng có dư nợ cuối + đến hạn theo kỳ hạn
(NH=Ngắn, TH=Trung) nhưng KHÔNG có đầu/vay/trả theo kỳ hạn.

Cách làm (giữ tổng per-ngân-hàng KHỚP group -> không regress thẻ đầu/vay/trả):
  - Mỗi dòng VAY (đơn vị, ngân hàng) hiện có -> tách thành các dòng con theo kỳ hạn.
  - Tỷ trọng kỳ hạn = dư nợ cuối per kỳ hạn (từ báo cáo NH).  (TC: Ngắn=NH sau cơ cấu,
    Trung=phần dư còn lại của group).
  - amount(dư cuối)/amount2(lãi)/du_dau/vay_them/tra_no = group × tỷ trọng (tổng khớp).
  - den_han = CHÍNH XÁC theo kỳ hạn (ma trận / cột 'Trả gốc' TC), không theo tỷ trọng.
  - dim2 = kỳ hạn. Ngân hàng không có báo cáo NH (vd 'Vay cá nhân') -> GIỮ NGUYÊN 1 dòng.
Chỉ TC/XVP/VFQN/AAG. Chạy SAU extract_sodutien_vay + extract_no_den_han + extract_lai_vay
(đọc dòng VAY đã có den_han/lãi để redistribute).

Chạy:  .venv/bin/python scripts/extract_vay_kyhan.py            # dry-run + đối soát
       .venv/bin/python scripts/extract_vay_kyhan.py --commit
"""
import argparse
import glob
import json
import os
import sys

sys.path.insert(0, "/home/sysadmin")
import cashflow_vay_extractor as ext            # noqa: E402
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers.common import be_bridge as bb      # noqa: E402

MASTER = {"ThinhCuong": "TC", "XanhVP": "XVP", "QuangNinh": "VFQN", "AN": "AAG"}
# dim2 phải là MÃ viết tắt NH/TH/DH — khớp FE (CashFlow.tsx TERM_LABEL/bankStackOption)
# và quy ước template 04_VAY "Kỳ hạn (NH/TH/DH)". FE hiển thị nhãn đầy đủ qua TERM_LABEL.
NGAN, TRUNG = "NH", "TH"


def bank_of_sheet(nm):
    s = nm.upper()
    if "BAB" in s: return "Bắc Á"
    if "TH-BIDV" in s: return "BIDV"
    if "TH-MB" in s or nm == "TH-MB AN": return "MB"
    if "VFQN" in s or "VPB" in s: return "VP"
    if "NH SAU CƠ CẤU" in s or "DƯ NỢ NH BIDV" in s or "TRUNG HẠN" in s: return "BIDV"
    return None


def nh_cuoi(ws):
    for r in ws.iter_rows(min_row=1, max_row=6, values_only=True):
        if any(isinstance(c, str) and "dư nợ còn phải trả" in c.lower() for c in r):
            nums = [v for v in r if isinstance(v, (int, float)) and abs(v) > 1e6]
            return nums[0] if nums else None
    return None


def th_cuoi(ws):
    for r in ws.iter_rows(min_row=1, max_row=8, values_only=True):
        if len(r) > 1 and isinstance(r[1], str) and "tổng vay" in r[1].lower():  # dòng 'TỔNG VAY ...'
            return r[10] if len(r) > 10 and isinstance(r[10], (int, float)) else None
    return None


def term_map(unit, f, month):
    """{bank: {term: {'cuoi': VND|None, 'den_han': tỷ}}} từ báo cáo ngân hàng."""
    wb = ext.load(f)
    tm = {}
    for nm in wb.sheetnames:
        if nm.startswith("foxz"):
            continue
        k = ext.sheet_kind(wb[nm])
        bank = bank_of_sheet(nm)
        if not bank:
            continue
        if k == "NH":
            dh = ext.matrix_month(wb[nm], month)
            tm.setdefault(bank, {})[NGAN] = {"cuoi": nh_cuoi(wb[nm]), "den_han": (dh or 0) / 1e9}
        elif k == "TH":
            dh = ext.matrix_month(wb[nm], month)
            tm.setdefault(bank, {})[TRUNG] = {"cuoi": th_cuoi(wb[nm]), "den_han": (dh or 0) / 1e9}
    if unit == "ThinhCuong":
        dh = 0.0
        for nm in ("Trung hạn 2022", "Trung hạn BIDV"):
            if nm in wb.sheetnames:
                v, _ = ext.wide_tra_goc(wb[nm], month)
                dh += (v or 0)
        tm.setdefault("BIDV", {})[TRUNG] = {"cuoi": None, "den_han": dh / 1e9}  # cuoi = residual
    wb.close()
    return tm


def split_row(row, tmap):
    """row (dict raw_rows) -> list dòng con theo kỳ hạn (giữ tổng)."""
    bank = row["dim1"]
    terms = tmap.get(bank)
    if not terms:                       # ví dụ 'Vay cá nhân' -> giữ nguyên
        return [row]
    amt = row["amount"] or 0.0          # dư cuối (tỷ)
    pl = json.loads(row["payload"] or "{}")
    # tỷ trọng theo dư nợ cuối per kỳ hạn (VND); TC/Trung = residual
    comp = {}
    for t, d in terms.items():
        comp[t] = d["cuoi"]
    if TRUNG in comp and comp[TRUNG] is None:               # TC: residual
        short = comp.get(NGAN) or 0.0
        comp[TRUNG] = max(0.0, amt * 1e9 - short)
    tot = sum(v for v in comp.values() if v) or 0.0
    out = []
    for t in terms:
        ratio = (comp[t] / tot) if tot > 0 else (1.0 / len(terms))
        npl = dict(pl)
        npl["du_dau_ky"] = (pl.get("du_dau_ky") or 0) * ratio
        npl["vay_them"] = (pl.get("vay_them") or 0) * ratio
        npl["tra_no"] = (pl.get("tra_no") or 0) * ratio
        npl["den_han"] = terms[t]["den_han"]               # CHÍNH XÁC per kỳ hạn
        nr = dict(row)
        nr["amount"] = amt * ratio
        nr["amount2"] = (row["amount2"] or 0.0) * ratio
        nr["dim2"] = t
        nr["payload"] = json.dumps(npl, ensure_ascii=False)
        out.append(nr)
    return out


def run(commit):
    db = bb.db.get_db()
    # gom file báo cáo NH theo (unit, month)
    files = {}
    for f in glob.glob(os.path.join(ext.DATA_DIR, "*.xlsx")):
        u, m = ext.unit_of(os.path.basename(f)), ext.month_of(os.path.basename(f))
        if u in MASTER and m:
            files[(u, m)] = f

    total_new = 0
    for (unit, month), f in sorted(files.items()):
        cty = MASTER[unit]
        period = f"2026-{month:02d}"
        rows = [dict(r) for r in db.execute(
            "SELECT * FROM raw_rows WHERE report_type='VAY' AND period_month=? AND cong_ty=?",
            (period, cty)).fetchall()]
        if not rows:
            continue
        tmap = term_map(unit, f, month)
        new_rows, ids = [], []
        for r in rows:
            ids.append(r["id"])
            new_rows.extend(split_row(r, tmap))
        # đối soát: tổng dư cuối + đến hạn phải giữ
        old_cuoi = sum(r["amount"] or 0 for r in rows)
        new_cuoi = sum(r["amount"] or 0 for r in new_rows)
        old_dh = sum(json.loads(r["payload"] or "{}").get("den_han") or 0 for r in rows)
        new_dh = sum(json.loads(r["payload"] or "{}").get("den_han") or 0 for r in new_rows)
        flag = "OK" if abs(old_cuoi - new_cuoi) < 0.01 else "⚠️CUOI-LỆCH"
        byterm = {}
        for r in new_rows:
            byterm[r["dim2"] or "—"] = byterm.get(r["dim2"] or "—", 0) + (json.loads(r["payload"] or "{}").get("den_han") or 0)
        print(f"{period} {cty}: {len(rows)}→{len(new_rows)} dòng | dư cuối {old_cuoi:.1f}→{new_cuoi:.1f} [{flag}] "
              f"| đến hạn {old_dh:.2f}→{new_dh:.2f} theo kỳ {({k: round(v,2) for k,v in byterm.items()})}")
        if commit:
            db.execute(f"DELETE FROM raw_rows WHERE id IN ({','.join('?'*len(ids))})", ids)
            db.executemany(
                "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,"
                "cost_center,period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                [(r["dataset_id"], r["report_type"], r["row_index"], r["ngay"], r["cong_ty"],
                  r["khoi"], r["cost_center"], r["period_month"], r["amount"], r["amount2"],
                  r["dim1"], r["dim2"], r["dim3"], r["payload"], r["source_file"]) for r in new_rows])
        total_new += len(new_rows)
    if commit:
        db.commit()
    print(f"\n{'ĐÃ GHI' if commit else 'DRY-RUN'}: {total_new} dòng con. "
          f"{'' if commit else 'Chạy --commit để ghi.'}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    run(ap.parse_args().commit)


if __name__ == "__main__":
    main()

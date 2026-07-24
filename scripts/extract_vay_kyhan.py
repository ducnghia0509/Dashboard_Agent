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


def _bank_of_bcth2(rec):
    """Dòng credit BCTH 2 -> mã ngân hàng (khớp dim1 VAY). Dò trong bank/name/code/tt."""
    t = " ".join(str(rec.get(k) or "") for k in ("bank", "name", "code", "tt")).upper()
    if "BAB" in t or "BẮC Á" in t or "BAC A" in t: return "Bắc Á"
    if "BIDV" in t: return "BIDV"
    if "VPB" in t or "VPBANK" in t: return "VP"
    if "-MB" in t or " MB" in t or "MB " in t: return "MB"
    if "VCB" in t or "VIETCOMBANK" in t: return "VCB"
    return None


def _is_lease(rec):
    """Cty cho thuê tài chính (Chailease, TK 341x) — KHÔNG phải ngân hàng, bỏ (tránh carry-forward nhầm)."""
    t = " ".join(str(rec.get(k) or "") for k in ("bank", "name", "code")).lower()
    return "cho thuê" in t or "chailease" in t or str(rec.get("code") or "").startswith("341")


def _term_of_bcth2(rec):
    t = str(rec.get("name") or "").lower()
    if "ngắn" in t or "ngan" in t: return NGAN
    if "trung" in t or "dài" in t or "dai" in t: return TRUNG   # gộp dài hạn vào TH (giữ dim2 NH/TH)
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
    """{bank: {term: {'cuoi','vay','tra','den_han' (tỷ)}}}. cuoi/vay/trả từ **BCTH 2** (per bank×kỳ hạn:
    Nợ cuối / Tổng số vay / Đã thanh toán) — nguồn per-bank chuẩn cho chart Đi vay/Trả nợ theo ngân hàng.
    den_han từ ma trận NH/TH (nợ đáo hạn trong tháng). Bỏ Chailease (leasing); carry-forward bank cho
    dòng con (BCTH2 gom facility dưới 1 dòng bank)."""
    wb = ext.load(f)
    tm = {}

    def _cell(b, t):
        return tm.setdefault(b, {}).setdefault(t, {"cuoi": 0.0, "vay": 0.0, "tra": 0.0, "den_han": 0.0, "den_han_next": 0.0})

    res, _ = ext.read_bcth2(wb, unit)
    last = None
    for r in (res.get("credit") or []):
        if _is_lease(r):
            continue
        b = _bank_of_bcth2(r) or last
        if _bank_of_bcth2(r):
            last = _bank_of_bcth2(r)
        if not b:
            continue
        d = _cell(b, _term_of_bcth2(r) or NGAN)
        d["cuoi"] += (r["no_cuoi_ky"] or 0) / 1e9
        d["vay"] += (r["tong_vay"] or 0) / 1e9
        d["tra"] += (r["da_thanh_toan"] or 0) / 1e9
    for nm in wb.sheetnames:                       # den_han: ma trận cột-tháng sheet NH/TH
        if nm.startswith("foxz"):
            continue
        k = ext.sheet_kind(wb[nm]); bank = bank_of_sheet(nm)
        term = NGAN if k == "NH" else TRUNG if k == "TH" else None
        if bank and term:
            ws = wb[nm]
            _cell(bank, term)["den_han"] += (ext.matrix_month(ws, month) or 0) / 1e9
            if month < 12:                              # #8: nợ gốc đáo hạn THÁNG TỚI = cột tháng M+1 (ma trận)
                _cell(bank, term)["den_han_next"] += (ext.matrix_month(ws, month + 1) or 0) / 1e9
    if unit == "ThinhCuong":                        # TC trung hạn: cột rộng 'Trả gốc T<mm>'
        for nm in ("Trung hạn 2022", "Trung hạn BIDV"):
            if nm in wb.sheetnames:
                v, _ = ext.wide_tra_goc(wb[nm], month)
                if v:
                    _cell("BIDV", TRUNG)["den_han"] += v / 1e9
    wb.close()
    return tm


def split_row(row, tmap):
    """row (per BANK) -> dòng con per (bank, kỳ hạn). Dư nợ (amount): giữ TỔNG từ SD TIỀN × tỷ trọng
    BCTH2 cuoi. Vay thêm/Trả/Đến hạn: BCTH2 **TRỰC TIẾP** per (bank,kỳ hạn) (không allocate theo tỷ trọng)."""
    bank = row["dim1"]
    terms = tmap.get(bank)
    if not terms:                       # bank ngoài báo cáo NH (vd 'Vay cá nhân') -> giữ nguyên 1 dòng
        return [row]
    amt = row["amount"] or 0.0          # dư cuối (tỷ) — nguồn SD TIỀN
    pl = json.loads(row["payload"] or "{}")
    comp = {t: (terms[t].get("cuoi") or 0.0) for t in terms}   # tỷ trọng dư nợ per kỳ hạn (từ BCTH2)
    tot = sum(comp.values()) or 0.0
    out = []
    for t in terms:
        ratio = (comp[t] / tot) if tot > 0 else (1.0 / len(terms))
        npl = dict(pl)
        npl["du_dau_ky"] = round((pl.get("du_dau_ky") or 0) * ratio, 9)
        npl["vay_them"] = round(terms[t].get("vay") or 0.0, 9)   # per (bank,kỳ hạn) từ BCTH2 — KHÔNG ratio
        npl["tra_no"] = round(terms[t].get("tra") or 0.0, 9)
        npl["den_han"] = round(terms[t].get("den_han") or 0.0, 9)
        npl["den_han_next"] = round(terms[t].get("den_han_next") or 0.0, 9)   # #8: đáo hạn tháng tới
        nr = dict(row)
        nr["amount"] = round(amt * ratio, 9)
        nr["amount2"] = round((row["amount2"] or 0.0) * ratio, 9)
        nr["dim2"] = t
        nr["payload"] = json.dumps(npl, ensure_ascii=False)
        out.append(nr)
    return out


def run(commit, only_period=None):
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
        if only_period and period != only_period:   # autofill scope 1 kỳ -> KHÔNG tách lại kỳ đã split (mangle)
            continue
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


def apply(only_period=None, commit=True):
    """Autofill hook: tách kỳ hạn (dim2 NH/TH) cho 1 kỳ. GIỮ NGUYÊN dòng '(Cả *)' & 'Vay cá nhân' (không có term_map)."""
    return run(commit, only_period)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    run(ap.parse_args().commit)


if __name__ == "__main__":
    main()

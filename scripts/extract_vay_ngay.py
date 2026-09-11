# -*- coding: utf-8 -*-
"""VAY theo NGÀY (dư nợ đầu/cuối, vay thêm, trả nợ, đến hạn) — hoàn thiện tab "Vay" ở chế độ Ngày
(mapping dongtienrealtime.xlsx dòng 5-53; 8 chỉ tiêu KHÔNG có ghi chú "Cần IT" = đã có ở tab Tháng,
nay hoàn thiện ở tab Ngày). Ghi report_type mới "VAY_D" (đã đăng ký ở
`repository._DAY_REPORT_TYPES` — cơ chế hậu tố `_D` DÙNG CHUNG với HQKD/CHIPHI/…).

Công thức mapping (đúng NGUYÊN VĂN các ô/cột, đọc từ file `baocaonganhang` — CÙNG NGUỒN đang phục
vụ tab Tháng, không phải file mới, xem `cashflow_vay_extractor.py`):
  · Hưng Thịnh: mapping ghi "Theo tháng (dùng mapping tháng)" -> KHÔNG diff theo ngày, mỗi ngày
    trong tháng lấy NGUYÊN vay_them/tra_no/den_han* mà `extract_vay_kyhan.term_map()` đã tính cho
    CẢ THÁNG (giống mọi ngày trong tháng đó).
  · Thịnh Cường (BIDV, kỳ hạn NH) — sheet "Dư nợ NH BIDV":
      vay_them = Σcột "số tiền vay" (ngày N) − Σcột (ngày N-1)
      tra_no   = Σcột "Gốc đã trả"  (ngày N) − Σcột (ngày N-1)
  · Xanh VP (VP, NH) — sheet "NH VPB":
      vay_them = Σcột "số tiền vay" (N) − (N-1)
      tra_no   = Σcột "dư nợ đến thời điểm hiện tại" (N-1) − (N)   [cột là SỐ DƯ GIẢM DẦN, đảo chiều]
    Xanh VP (VP, TH) — sheet "TH VPB", dòng tổng (ô "Số tiền vay"/"Dư nợ vay" dòng 5):
      vay_them = ô(N) − ô(N-1); tra_no = ô(N) − ô(N-1)   [ĐÚNG NGUYÊN VĂN mapping — xem cảnh báo dưới]
    Xanh VP (Bắc Á, TH) — sheet "TH BAB", cùng công thức TH VPB.
  · VFQN (VP, NH) — sheet "NH VFQN": như Xanh VP (VP, NH) ở trên (cùng bank VPBank, đảo chiều tra_no).
  · An: chưa có nguồn — bỏ qua.

⚠️ CẢNH BÁO CẦN ĐỐI CHIẾU TAY TRƯỚC KHI TIN SỐ: mapping dùng ô K5 ("Dư nợ vay" — số dư CÒN LẠI,
giảm dần khi trả nợ) với diff CHIỀU THƯỜNG (N − N-1) cho "trả nợ" của TH VPB/TH BAB — trong khi
cùng loại cột ("dư nợ hiện tại") ở sheet NH VPB/NH VFQN lại dùng diff ĐẢO CHIỀU. Đã cài ĐÚNG NGUYÊN
VĂN mapping (không tự sửa theo suy luận riêng); nếu số ra âm/dương ngược trực giác ở TH VPB/TH BAB,
đối chiếu lại với KTV trước khi tin dùng cho báo cáo.

Cần "hôm qua" để diff: lưu kèm trong payload mỗi dòng VAY_D 2 field ẩn (không hiển thị FE)
`_cum_vay`/`_cum_tra` = giá trị LŨY KẾ THÔ đọc được hôm đó (trước khi diff) — hôm sau tự tra lại
dòng VAY_D của (cty,bank,term) hôm qua để diff tiếp, khỏi cần bảng phụ. Ngày đầu tiên chạy (không
có "hôm qua") -> vay_them/tra_no = 0, dư nợ cuối NGÀY = SEED từ dòng VAY (tháng) hiện có.

Chạy: .venv/bin/python scripts/extract_vay_ngay.py [--commit] [--ngay yyyy-mm-dd]
"""
import argparse
import datetime
import json
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..", "..")))
sys.path.insert(0, _HERE)
import cashflow_vay_extractor as ext          # noqa: E402
import extract_vay_kyhan as evk               # noqa: E402
sys.path.insert(0, os.path.dirname(_HERE))
from servers.common import be_bridge as bb    # noqa: E402

MASTER = evk.MASTER   # {"ThinhCuong":"TC","XanhVP":"XVP","QuangNinh":"VFQN","AN":"AAG"}
YEAR = 2026


def _header_row_idx(rows, kw, max_scan=10):
    for i, r in enumerate(rows[:max_scan]):
        if ext.find_col(r, kw) is not None:
            return i, ext.find_col(r, kw)
    return None, None


def col_sum(ws, kw):
    """Σ cột có header khớp `kw` (dò 10 dòng đầu), CHỈ cộng dòng SAU header (tránh ô tổng trang trí
    nằm phía TRÊN header, vd 'NH VPB'!G2 — một ô tổng khác, không phải dòng dữ liệu)."""
    rows = ext.rows_of(ws, 400)
    i, j = _header_row_idx(rows, kw)
    if i is None:
        return None
    total = 0.0
    for r in rows[i + 1:]:
        v = ext.num(ext.get(r, j))
        if v:
            total += v
    return total


def cell_at(ws, kw, target_row_1based):
    """Giá trị tại DÒNG `target_row_1based`, CỘT có header khớp `kw`."""
    rows = ext.rows_of(ws, max(target_row_1based, 10))
    _, j = _header_row_idx(rows, kw)
    if j is None or target_row_1based - 1 >= len(rows):
        return None
    return ext.num(ext.get(rows[target_row_1based - 1], j))


def cum_today(wb, cty):
    """-> {(bank, term): (cum_vay, cum_tra, tra_reversed)} — giá trị LŨY KẾ thô đọc hôm nay,
    CHƯA diff. `tra_reversed`=True nghĩa là cột trả nợ là SỐ DƯ GIẢM DẦN (diff N-1 trừ N)."""
    # LƯU Ý: `ext.find_col`/`s()` chỉ lowercase, KHÔNG bỏ dấu — từ khoá dò cột phải giữ NGUYÊN dấu
    # tiếng Việt khớp đúng header thật trong file nguồn.
    out = {}
    if cty == "TC":
        if "Dư nợ NH BIDV" in wb.sheetnames:
            ws = wb["Dư nợ NH BIDV"]
            out[("BIDV", "NH")] = (col_sum(ws, "số tiền vay"), col_sum(ws, "gốc đã trả"), False)
    elif cty == "VFQN":
        if "NH VFQN" in wb.sheetnames:
            ws = wb["NH VFQN"]
            out[("VP", "NH")] = (col_sum(ws, "số tiền vay"),
                                  col_sum(ws, "đến thời điểm hiện tại"), True)
    elif cty == "XVP":
        if "NH VPB" in wb.sheetnames:
            ws = wb["NH VPB"]
            out[("VP", "NH")] = (col_sum(ws, "số tiền vay"),
                                  col_sum(ws, "đến thời điểm hiện tại"), True)
        if "TH VPB" in wb.sheetnames:
            ws = wb["TH VPB"]
            out[("VP", "TH")] = (cell_at(ws, "số tiền vay", 5), cell_at(ws, "dư nợ vay", 5), False)
        if "TH BAB" in wb.sheetnames:
            ws = wb["TH BAB"]
            out[("Bắc Á", "TH")] = (cell_at(ws, "số tiền vay", 5), cell_at(ws, "dư nợ vay", 5), False)
    return out


def _yesterday_row(db, cty, bank, term, ngay_hom_qua):
    r = db.execute(
        "SELECT amount, payload FROM raw_rows WHERE report_type='VAY_D' AND cong_ty=? AND dim1=? "
        "AND COALESCE(dim2,'')=? AND ngay=?",
        (cty, bank, term or "", ngay_hom_qua)).fetchone()
    if not r:
        return None
    try:
        pl = json.loads(r["payload"] or "{}")
    except Exception:
        pl = {}
    return {"amount": r["amount"], "payload": pl}


def _month_seed(db, cty, bank, term, period):
    """Seed NGÀY ĐẦU TIÊN (không có 'hôm qua'): lấy dư nợ cuối kỳ + du_dau_ky từ dòng VAY (tháng)
    hiện có của đúng (cty,bank,term) — tránh bắt đầu từ 0 một cách vô nghĩa."""
    r = db.execute(
        "SELECT amount, payload FROM raw_rows WHERE report_type='VAY' AND cong_ty=? AND dim1=? "
        "AND COALESCE(dim2,'')=? AND period_month=?",
        (cty, bank, term or "", period)).fetchone()
    if not r:
        return 0.0, 0.0
    try:
        pl = json.loads(r["payload"] or "{}")
    except Exception:
        pl = {}
    return (r["amount"] or 0.0), (pl.get("du_dau_ky") or 0.0)


def _dataset_id_for(db, cty, period):
    row = db.execute(
        "SELECT dataset_id FROM raw_rows WHERE report_type='VAY' AND cong_ty=? AND period_month=? "
        "LIMIT 1", (cty, period)).fetchone()
    return row["dataset_id"] if row else None


def run(commit, ngay=None):
    ngay = ngay or datetime.date.today().isoformat()
    hom_qua = (datetime.date.fromisoformat(ngay) - datetime.timedelta(days=1)).isoformat()
    period = ngay[:7]
    month = int(period[5:7])
    db = bb.db.get_db()

    files = {}
    for f in ext.files_can_doc():
        u, m = ext.unit_of(os.path.basename(f)), ext.month_of(os.path.basename(f))
        if u in MASTER and m == month:
            files[MASTER[u]] = f

    out_rows, skipped = [], []

    # Hưng Thịnh: theo tháng (mapping tháng) — copy nguyên vay_them/tra_no/den_han* của term_map(),
    # KHÔNG diff ngày. File HT không nằm trong MASTER (dict evk.MASTER không có "HungThinh") nên
    # xử lý riêng, đọc trực tiếp file HT + term_map "HungThinh".
    ht_file = next((f for f in ext.files_can_doc()
                    if ext.unit_of(os.path.basename(f)) == "HungThinh"
                    and ext.month_of(os.path.basename(f)) == month), None)
    if ht_file:
        tm = evk.term_map("HungThinh", ht_file, month)
        ds = _dataset_id_for(db, "HT", period)
        if ds is None:
            skipped.append({"cong_ty": "HT", "error": f"chua_co_dataset_VAY_ky_{period}"})
        else:
            for bank, terms in tm.items():
                for term, v in terms.items():
                    seed_cuoi, seed_dau = _month_seed(db, "HT", bank, term, period)
                    row = {"cong_ty": "HT", "dim1": bank, "dim2": term, "dataset_id": ds,
                           "amount": seed_cuoi, "du_dau_ky": seed_dau,
                           "vay_them": v.get("vay", 0.0), "tra_no": v.get("tra", 0.0),
                           "den_han": v.get("den_han", 0.0), "den_han_next": v.get("den_han_next", 0.0),
                           "den_han_next2": v.get("den_han_next2", 0.0),
                           "cum_vay": None, "cum_tra": None, "source": ht_file}
                    out_rows.append(row)

    for cty, f in files.items():
        wb = ext.load(f)
        tm = evk.term_map({"TC": "ThinhCuong", "XVP": "XanhVP", "VFQN": "QuangNinh"}[cty], f, month)
        cum = cum_today(wb, cty)
        wb.close()
        ds = _dataset_id_for(db, cty, period)
        if ds is None:
            skipped.append({"cong_ty": cty, "error": f"chua_co_dataset_VAY_ky_{period}"})
            continue
        for (bank, term), (cum_vay, cum_tra, tra_rev) in cum.items():
            terms_month = tm.get(bank, {}).get(term, {})
            y = _yesterday_row(db, cty, bank, term, hom_qua)
            if y is None or cum_vay is None or cum_tra is None:
                seed_cuoi, seed_dau = _month_seed(db, cty, bank, term, period)
                du_dau, vay_them, tra_no = seed_dau, 0.0, 0.0
                du_cuoi = seed_cuoi if (y is None) else (du_dau + vay_them - tra_no)
            else:
                y_pl = y["payload"]
                y_cum_vay, y_cum_tra = y_pl.get("cum_vay"), y_pl.get("cum_tra")
                vay_them = ((cum_vay - y_cum_vay) / 1e9) if y_cum_vay is not None else 0.0
                tra_no = (((cum_tra - y_cum_tra) if not tra_rev else (y_cum_tra - cum_tra)) / 1e9) \
                    if y_cum_tra is not None else 0.0
                du_dau = (y["amount"] or 0.0)
                du_cuoi = du_dau + vay_them - tra_no
            out_rows.append({"cong_ty": cty, "dim1": bank, "dim2": term, "dataset_id": ds,
                             "amount": round(du_cuoi, 9), "du_dau_ky": round(du_dau, 9),
                             "vay_them": round(vay_them, 9), "tra_no": round(tra_no, 9),
                             "den_han": terms_month.get("den_han", 0.0),
                             "den_han_next": terms_month.get("den_han_next", 0.0),
                             "den_han_next2": terms_month.get("den_han_next2", 0.0),
                             "cum_vay": cum_vay, "cum_tra": cum_tra, "source": f})

    if commit:
        for r in out_rows:
            payload = {"du_dau_ky": r["du_dau_ky"], "vay_them": r["vay_them"], "tra_no": r["tra_no"],
                       "den_han": r["den_han"], "den_han_next": r["den_han_next"],
                       "den_han_next2": r["den_han_next2"],
                       "cum_vay": r["cum_vay"], "cum_tra": r["cum_tra"]}
            db.execute("DELETE FROM raw_rows WHERE report_type='VAY_D' AND ngay=? AND cong_ty=? "
                       "AND dim1=? AND COALESCE(dim2,'')=?",
                       (ngay, r["cong_ty"], r["dim1"], r["dim2"] or ""))
            db.execute(
                "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,cost_center,"
                "period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (r["dataset_id"], "VAY_D", 8000000, ngay, r["cong_ty"], None, None, period,
                 r["amount"], None, r["dim1"], r["dim2"], None,
                 json.dumps(payload, ensure_ascii=False), r["source"]))
        db.commit()
    return {"ngay": ngay, "ok": out_rows, "skipped": skipped}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--ngay", default=None)
    a = ap.parse_args()
    out = run(a.commit, a.ngay)
    for r in out["ok"]:
        print(f"{out['ngay']} {r['cong_ty']:5s} {r['dim1']:6s} {r['dim2'] or '—':3s} "
              f"đầu={r['du_dau_ky']:.3f} vay={r['vay_them']:.3f} trả={r['tra_no']:.3f} "
              f"cuối={r['amount']:.3f} đến_hạn={r['den_han']:.3f}/{r['den_han_next']:.3f}/{r['den_han_next2']:.3f}")
    for s in out["skipped"]:
        print("BỎ QUA:", s)
    print(f"\n{'ĐÃ GHI' if a.commit else 'DRY-RUN'}: {len(out['ok'])} dòng, {len(out['skipped'])} bỏ qua.")


if __name__ == "__main__":
    main()

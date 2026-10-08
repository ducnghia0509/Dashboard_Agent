"""Soát CÔNG THỨC trên chính số dashboard hiển thị — mỗi khối × tháng và mỗi khối × ngày.

VÌ SAO (08/10/2026): báo cáo ngày Xe tải T9 hiện LNST 2,63 tỷ trong khi DT 52,53 + DTTC&TN khác
1,27 − Chi phí 52,48 = 1,31. Dòng T500 "LỢI NHUẬN SAU THUẾ TNDN" bị deriver ghi thêm lần thứ hai
bên cạnh dòng "Lợi nhuận sau thuế" chuẩn, thẻ LNST ở Tổng quan (PNLT ILIKE '%lợi nhuận%sau thu%')
cộng cả hai. HQKD 1112 (màn Hiệu quả KD) vẫn đúng 1,31 — tức HAI MÀN RA HAI SỐ cho cùng một chỉ
tiêu suốt 3 tuần mà không ai biết. Cùng lỗi đã xảy ra ở bản THÁNG (23/09) qua một đường code khác.
Cả hai lần đều là phép cộng trừ đơn giản phát hiện được, nên soát bằng phép cộng trừ.

Hai phép soát:
  C1  LNST hai màn phải bằng nhau: Σ PNLT "lợi nhuận…sau thuế" (Tổng quan) == HQKD 1112 (Hiệu quả KD).
  C2  Công thức kế toán: Doanh thu + DT tài chính & TN khác − Chi phí (1047) == LNTT (1112 + 1111).

TÍNH BẰNG CHÍNH HÀM CỦA BACKEND (`metrics.flow_by_khoi` + `grain_scope`), không tự viết SQL: chọn
dòng của dashboard có nhiều tầng (mỗi file đóng góp 1 lần — anchor ưu tiên hơn cost center; báo cáo
ngày chỉ lấy FILE MỚI NHẤT mỗi lát; trừ `hidden_files`; DTHU ưu tiên hơn HQKD 1000). Viết lại bằng
SQL là tự sinh ra báo động giả — đúng loại sai đã gặp ở bộ dò cộng đôi QTVH 09/09. Vì vậy script
PHẢI chạy bằng python + cwd của một checkout tc-admin-api:

  cd <tc-admin-api> && DATABASE_URL=... .venv/bin/python <Dashboard_Agent>/scripts/soat_cong_thuc.py \
      [--ky 2026-09 --ky 2026-10] [--grain day|month|all] [--json]

CHỈ BÁO, KHÔNG SỬA: sửa tự động xoá mất dấu vết nguồn lỗi. Ngoại lệ hợp lệ (khối mà công thức
không khép vì lý do kế toán đã biết) khai ở `soat_cong_thuc_ngoai_le.json` cạnh file này, mỗi mục
BẮT BUỘC có `ly_do` — không có lý do thì không phải ngoại lệ, là lỗi chưa điều tra.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.getcwd())

from app import metrics as M                                   # noqa: E402
from app.database.session import get_db                         # noqa: E402
from app.metrics import repository as mr                        # noqa: E402

NGOAI_LE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "soat_cong_thuc_ngoai_le.json")

# Dung sai (đơn vị amount = TỶ đồng): 1 triệu, hoặc 0,1% của quy mô lát (max |DT|, |CP|). Số trên
# dashboard đã qua làm tròn ở nguồn từng dòng nên lệch vài trăm nghìn ở lát to là bình thường.
TOL_TUYET_DOI = 0.001
TOL_TUONG_DOI = 0.001


def _tol(*vals):
    return max(TOL_TUYET_DOI, TOL_TUONG_DOI * max((abs(v) for v in vals if v is not None), default=0))


def _cau_hinh():
    try:
        with open(NGOAI_LE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def _khop_ngoai_le(nl, check, khoi, grain, ky):
    """Ngoại lệ khớp khi cùng phép soát + cùng khối, và (nếu khai) cùng grain / ky nằm trong
    [tu, den]. `ky` so chuỗi: '2026-09' (tháng) hoặc '2026-09-18' (ngày) đều so được với mốc
    'YYYY-MM' / 'YYYY-MM-DD' vì cùng thứ tự từ điển."""
    for e in nl:
        if e.get("check") not in (check, "*") or e.get("khoi") != khoi:
            continue
        if e.get("grain") and e["grain"] != grain:
            continue
        if e.get("tu") and ky < e["tu"]:
            continue
        if e.get("den") and ky[:len(e["den"])] > e["den"]:
            continue
        return e
    return None


def _datasets(kys):
    rows = get_db().execute(
        "SELECT id, period FROM datasets WHERE kind='month' AND period IS NOT NULL "
        "AND COALESCE(hidden,0)=0 ORDER BY period").fetchall()
    return [(r["id"], r["period"]) for r in rows if not kys or r["period"] in kys]


def _ngay_cua(ds):
    rows = get_db().execute(
        "SELECT DISTINCT ngay FROM raw_rows WHERE dataset_id=? AND ngay IS NOT NULL "
        "AND report_type IN ('HQKD_D','PNLT_D') ORDER BY ngay", (ds,)).fetchall()
    return [str(r["ngay"])[:10] for r in rows]


def _ngoai_the(ds, frm, to, khoan):
    """Σ theo khối các khoản CÓ trong LNTT của nguồn nhưng CỐ Ý không thuộc thẻ nào (khai ở
    `khoan_ngoai_the`). Cộng vào vế công thức thay vì miễn soát cả khối: nhờ vậy mọi lệch KHÁC của
    khối đó vẫn bị bắt."""
    out = {}
    for k in khoan:
        v = M.flow_by_khoi(ds, "PNLT", frm, to, dim1_ilike=k["dim1_ilike"]).get(k["khoi"])
        if v:
            out[k["khoi"]] = out.get(k["khoi"], 0.0) + v
    return out


def _lat(ds, frm, to):
    """Số 4 thẻ Tổng quan + 1112/1111 theo KHỐI cho 1 lát — y hệt overview.build_overview."""
    fk = M.flow_by_khoi
    dthu_has = "DTHU" in M.available_reports(ds)
    dt = fk(ds, "DTHU", frm, to) if dthu_has else fk(ds, "HQKD", frm, to, dim1=mr.HQKD_REVENUE)
    out = {
        "dt": dt,
        "fin": fk(ds, "PNLT", frm, to, dim1_ilike="%doanh thu%tài chính%"),
        "tnk": fk(ds, "PNLT", frm, to, dim1_ilike="%thu nhập khác%"),
        "cp": fk(ds, "HQKD", frm, to, dim1=mr.HQKD_COST),
        "lnst": fk(ds, "PNLT", frm, to, dim1_ilike=mr.LNST_ILIKE),
        "m1112": fk(ds, "HQKD", frm, to, dim1=mr.HQKD_PROFIT_AT),
        "m1111": fk(ds, "HQKD", frm, to, dim1=mr.HQKD_TAX_TNDN),
    }
    khois = set().union(*(v.keys() for v in out.values()))
    return {k: {n: out[n].get(k) for n in out} for k in khois}


def _soat_lat(grain, ky, ds, frm, to, nl, khoan):
    loi, bo_qua = [], []
    them = _ngoai_the(ds, frm, to, khoan) if khoan else {}
    for khoi, v in sorted(_lat(ds, frm, to).items()):
        dt, cp, lnst, m1112 = v["dt"], v["cp"], v["lnst"], v["m1112"]
        fin = (v["fin"] or 0) + (v["tnk"] or 0) + them.get(khoi, 0.0)
        lntt = None if m1112 is None else m1112 + (v["m1111"] or 0)
        tol = _tol(dt, cp)
        phat = []
        # C1 chỉ soát khi CẢ HAI phía có số: thiếu một phía là chuyện "nguồn chưa có chỉ tiêu",
        # không phải hai màn mâu thuẫn nhau.
        if lnst is not None and m1112 is not None and abs(lnst - m1112) > tol:
            phat.append({"check": "C1", "lnst_tong_quan": lnst, "lnst_hqkd_1112": m1112,
                         "chenh": lnst - m1112,
                         "gap_doi": abs(m1112) > tol and abs(lnst - 2 * m1112) <= tol})
        if dt is not None and cp is not None and lntt is not None:
            ct = dt + fin - cp
            if abs(ct - lntt) > tol:
                phat.append({"check": "C2", "doanh_thu": dt, "dttc_tnkhac": fin, "chi_phi": cp,
                             "cong_thuc": ct, "lntt": lntt, "chenh": ct - lntt})
        for p in phat:
            rec = {"grain": grain, "ky": ky, "khoi": khoi, **p}
            e = _khop_ngoai_le(nl, p["check"], khoi, grain, ky)
            if e:
                rec["ngoai_le"] = e.get("ly_do")
                bo_qua.append(rec)
            else:
                loi.append(rec)
    return loi, bo_qua


def soat(kys=None, grain="all"):
    cfg = _cau_hinh()
    nl, khoan = cfg.get("ngoai_le", []), cfg.get("khoan_ngoai_the", [])
    loi, bo_qua, so_lat = [], [], 0
    for ds, period in _datasets(kys):
        if grain in ("all", "month"):
            with M.grain_scope("month"):
                a, b = _soat_lat("month", period, ds, None, None, nl, khoan)
            loi += a; bo_qua += b; so_lat += 1
        if grain in ("all", "day"):
            with M.grain_scope("day"):
                for d in _ngay_cua(ds):
                    a, b = _soat_lat("day", d, ds, d, d, nl, khoan)
                    loi += a; bo_qua += b; so_lat += 1
    return {"so_lat": so_lat, "so_loi": len(loi), "so_ngoai_le": len(bo_qua),
            "loi": loi, "ngoai_le": bo_qua}


def _f(x):
    return "—" if x is None else f"{x:,.4f}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ky", action="append", help="kỳ YYYY-MM (lặp lại được); bỏ trống = mọi kỳ")
    ap.add_argument("--grain", choices=("all", "month", "day"), default="all")
    ap.add_argument("--json", action="store_true", help="in JSON (cho cron)")
    args = ap.parse_args()
    kq = soat(args.ky, args.grain)
    if args.json:
        print(json.dumps(kq, ensure_ascii=False, default=str))
        return 0
    print(f"{kq['so_lat']} lát kỳ | {kq['so_loi']} lệch | {kq['so_ngoai_le']} lệch đã khai ngoại lệ")
    for r in kq["loi"]:
        if r["check"] == "C1":
            print(f"  C1 {r['grain']:5} {r['ky']:10} {r['khoi'][:34]:34} LNST Tổng quan {_f(r['lnst_tong_quan'])}"
                  f" ≠ 1112 {_f(r['lnst_hqkd_1112'])}{'  (GẤP ĐÔI)' if r['gap_doi'] else ''}")
        else:
            print(f"  C2 {r['grain']:5} {r['ky']:10} {r['khoi'][:34]:34} DT {_f(r['doanh_thu'])}"
                  f" + TC {_f(r['dttc_tnkhac'])} − CP {_f(r['chi_phi'])} = {_f(r['cong_thuc'])}"
                  f" ≠ LNTT {_f(r['lntt'])} (lệch {_f(r['chenh'])})")
    return 1 if kq["so_loi"] else 0


if __name__ == "__main__":
    sys.exit(main())

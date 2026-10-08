# -*- coding: utf-8 -*-
"""SỐ DƯ THÁNG DỰ PHÒNG từ báo cáo NGÀY CUỐI THÁNG — khi BCTC tháng chưa có (hoặc sheet chép kỳ khác).

VÌ SAO (08/10/2026): file BCTC tháng 9 của Trạm sạc là bản T8 chép lại — chỉ BCHQKD lên T09, CĐKT/
CĐPS/sổ công nợ/biểu khấu hao vẫn là số T8 ("Từ 01/08/2026 Đến 31/08/2026"). autofill nay bỏ các
sheet lệch kỳ (`skip_lech_ky`), nên các màn số dư tháng 9 trống: bảng "Biến động phải thu (khoản mục
CĐKT)" cuối kỳ 0, phải trả 0. Trong khi đó file báo cáo NGÀY 30/09 (B.3.TC.TCKT.D.20260930) có ĐỦ
bảng cân đối tại 30/09 — mã 131 = 9.927.254.609, khớp từng đồng với file tuổi nợ T9 — và đã nằm sẵn
trong DB dưới dạng *_D. CĐKT tại ngày cuối tháng CHÍNH LÀ CĐKT tháng.

LÀM GÌ: với mỗi khối khai trong `_KHOI`, mỗi loại số dư trong `_RT`: nếu khối/kỳ CHƯA có dòng tháng
nào của loại đó (ngoài chính bản dự phòng) và ĐÃ có dòng *_D của NGÀY CUỐI THÁNG -> chép sang loại
tháng, gắn cờ `du_phong_ngay_cuoi_thang`.
  - Cuối kỳ = số ngày cuối tháng (amount giữ nguyên).
  - Đầu kỳ (payload du_dau) = cuối kỳ THÁNG TRƯỚC của cùng đối tượng/khoản mục, không có thì None.
    Bản ngày ghi du_dau = số dư NGÀY HÔM TRƯỚC — đem lên tháng là sai nghĩa.
  - Phát sinh (ps_tang/ps_giam/nhap/xuat/khau_hao_ky) = None: bản ngày chỉ phủ MỘT ngày, Σ phát sinh
    các ngày của Trạm sạc T09 không khép với số dư (lệch 1,89 tỷ) nên không cộng dồn được.
  - TSNV amount2 (số ĐẦU NĂM) lấy từ CĐKT tháng trước (giống nhau mọi tháng trong năm).
  - Bản tháng THẬT nạp vào (template_filler.import_filled) tự xoá dòng mang cờ -> không cộng đôi.

CHỈ BẬT cho khối đã đối chiếu: Trạm sạc. Showroom/XDV có CĐKT ngày nhưng CHƯA CÂN (xem
derive_sodu_tudong_ngay.py) — đừng thêm vào `_KHOI` khi màn TS-NV của họ còn ô "LỆCH".

  DATABASE_URL=... .venv/bin/python scripts/du_phong_sodu_ngay_cuoi_thang.py --period 2026-09 [--write]
"""
import argparse
import calendar
import json
import os

import psycopg

DB_URL = (os.environ.get("DATABASE_URL") or os.environ.get("TC_DATABASE_URL")
          or "postgresql://tc:tc@localhost:5433/tc_dashboard")

_KHOI = ("Khối KD Trạm sạc Vgreen",)
# PTHU đứng cuối: Trạm sạc đã có bản dự phòng từ file tuổi nợ (tách được dư Nợ/dư Có, xem
# derive_congno_tuoino._ghi_pthu_du_phong) -> kỳ đó có PTHU rồi thì ở đây tự bỏ qua.
_RT = ("TSNV", "BS", "TS", "HH", "THUE", "PTRA", "PTHU")
_CO = '%"du_phong_ngay_cuoi_thang": true%'
_PS = ("ps_tang", "ps_giam", "nhap", "xuat", "khau_hao_ky")


def _ky_truoc(period):
    y, m = int(period[:4]), int(period[5:7])
    return f"{y - (m == 1)}-{(m - 2) % 12 + 1:02d}"


def _ngay_cuoi(period):
    y, m = int(period[:4]), int(period[5:7])
    return f"{period}-{calendar.monthrange(y, m)[1]:02d}"


def _pl(p):
    return json.loads(p) if isinstance(p, str) else dict(p or {})


def _khoa(dim1, dim2, pl):
    """Khoá nối cùng đối tượng/khoản mục giữa hai kỳ: mã đối tượng (công nợ) / mã số CĐKT / tên."""
    return pl.get("ma_dt") or pl.get("ma_so") or pl.get("tk") or f"{dim1}|{dim2}"


def dung(cur, period, write=False):
    ngay, truoc = _ngay_cuoi(period), _ky_truoc(period)
    kq = []
    for khoi in _KHOI:
        for rt in _RT:
            if write:
                cur.execute("DELETE FROM raw_rows WHERE report_type=%s AND period_month=%s AND khoi=%s "
                            "AND payload LIKE %s", (rt, period, khoi, _CO))
            cur.execute("SELECT 1 FROM raw_rows WHERE report_type=%s AND period_month=%s AND khoi=%s "
                        "AND payload NOT LIKE %s LIMIT 1", (rt, period, khoi, _CO))
            if cur.fetchone():
                kq.append({"khoi": khoi, "rt": rt, "ghi": False, "ly_do": "đã có bản tháng"})
                continue
            cur.execute("SELECT dataset_id, cong_ty, cost_center, dim1, dim2, dim3, amount, amount2, "
                        "payload, source_file FROM raw_rows WHERE report_type=%s AND period_month=%s "
                        "AND khoi=%s AND ngay=%s ORDER BY row_index", (rt + "_D", period, khoi, ngay))
            ngay_rows = cur.fetchall()
            if not ngay_rows:
                kq.append({"khoi": khoi, "rt": rt, "ghi": False, "ly_do": f"chưa có {rt}_D ngày {ngay}"})
                continue
            cur.execute("SELECT dim1, dim2, amount, amount2, payload FROM raw_rows WHERE report_type=%s "
                        "AND period_month=%s AND khoi=%s", (rt, truoc, khoi))
            dau = {}
            for d1, d2, amt, amt2, pl in cur.fetchall():
                k = _khoa(d1, d2, _pl(pl))
                e = dau.setdefault(k, [0.0, amt2])
                e[0] += float(amt or 0)
            tong = 0.0
            for i, (ds, cty, cc, d1, d2, d3, amt, amt2, pl, src) in enumerate(ngay_rows):
                # Bản ngày ghi dim RỖNG là '' còn bản tháng là NULL — debt._cdkt_accs nhận dòng LEAF
                # CĐKT bằng `dim1 is None`; giữ '' thì mọi khoản mục con bị bỏ, cả mục 130 dồn vào
                # "Phải thu khác" (32,35 tỷ Trạm sạc T09).
                d1, d2, d3, cc = d1 or None, d2 or None, d3 or None, cc or None
                p = _pl(pl)
                p.pop("grain", None)
                for f in _PS:
                    if f in p:
                        p[f] = None
                if "du_dau" in p:
                    p["du_dau"] = dau[_khoa(d1, d2, p)][0] if _khoa(d1, d2, p) in dau else None
                p["du_phong_ngay_cuoi_thang"] = True
                p["nguon_ngay"] = f"{src} @ {ngay}"
                if rt == "TSNV":   # số ĐẦU NĂM: bản ngày không có, lấy từ CĐKT tháng trước
                    amt2 = (dau.get(_khoa(d1, d2, p)) or [None, None])[1]
                tong += float(amt or 0)
                if write:
                    cur.execute(
                        "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
                        "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, source_file) "
                        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                        (ds, rt, 7700000 + i, ngay, cty, khoi, cc, period, amt, amt2, d1, d2, d3,
                         json.dumps(p, ensure_ascii=False), f"{src}::du_phong_thang"))
            kq.append({"khoi": khoi, "rt": rt, "ghi": write, "dong": len(ngay_rows),
                       "tong_ty": round(tong, 9), "dau_ky_tu": truoc if dau else None})
    return kq


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", required=True, action="append")
    ap.add_argument("--write", action="store_true", help="ghi DB (mặc định dry-run)")
    a = ap.parse_args()
    with psycopg.connect(DB_URL) as conn:
        cur = conn.cursor()
        out = {p: dung(cur, p, a.write) for p in a.period}
        if a.write:
            conn.commit()
    print(json.dumps(out, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()

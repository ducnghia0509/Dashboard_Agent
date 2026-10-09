#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: TUỔI NỢ PHẢI THU THEO NGÀY (`PTHU_TUOINO_D`) — kéo file `.D.YYYYMMDD.` rồi nạp.

VÌ SAO (09/10/2026, user chốt "theo ngày thì nên có cron để kéo"): từ 05/10 kế toán XVP, HTX Xanh
Vĩnh Phúc, HTX Xanh Tuyên Quang phát hành tuổi nợ THEO NGÀY (`B.6.<PN>.D.20261005.Baocaotuoino
phaithu.xlsx`, mỗi ngày một file) mà không job nào kéo: các cron QTVH cố ý bỏ thư mục `baocaotuoino`
(xem docstring `cron_xvp_daily.py`), còn bản tháng thì kéo tay. Kết quả: 10 file đã nằm trên nguồn,
cả hai DB không có dòng nào.

CHỌN FILE: mọi entry trong thư mục tuổi nợ (`report_type` chứa 'tuoino' — gồm cả `baocaotuoinongay`
khi máy gửi tách thư mục ngày) có mã kỳ ĐỦ NGÀY trong tên, `.D.20261005` hay `.M.20261005`/`M20261006`
(09/10/2026: sáu đơn vị ghi file ngày bằng `.M.` — xem `derive_congno_tuoino.ngay_trong_ten`), nằm
trong `CUA_SO_NGAY` ngày gần nhất. Tên sai quy ước (`.M.2026010`, `.M.202610.07`) không có ngày nên
không được chọn — phải báo kế toán sửa tên, không đoán.
Kéo lại cả cửa sổ chứ không chỉ hôm nay vì kế toán hay sửa lại file vài ngày trước. File ngày kiểu
cũ chỉ mang tháng (`.D.202608`, Showroom) KHÔNG thuộc job này — không có ngày để xếp.

NẠP: `derive_congno_tuoino.derive(path, period, write=True)` — idempotent theo source_file (xoá rồi
ghi lại), kỳ = tháng của ngày trong tên file. File không phải tuổi nợ (XDV `B.2…Baocaocongnophaithu`,
đối soát công nợ VF) deriver tự bỏ qua (`skip`).

`--pull` mới kéo. Quy ước crontab giống `cron_vay_ngay.py`: lượt PROD kéo, lượt TEST chạy sau vài
phút KHÔNG kéo (dùng chung thư mục received_reports).

Chạy: .venv/bin/python cron/cron_tuoino_ngay.py [--env test|prod] [--pull] [--ngay 10]
"""
import argparse
import os
import sys
from datetime import date, datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))
import cron_qtvh_core as core  # noqa: E402
import pull_nguon  # noqa: E402

VN = timezone(timedelta(hours=7))
CUA_SO_NGAY = 10


def log(msg):
    print(f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}", flush=True)


def _chon(e, tu: date, den: date, ngay_trong_ten) -> bool:
    if "tuoino" not in (e.get("report_type") or "").lower():
        return False
    d = ngay_trong_ten(e.get("fileName") or "")
    return d is not None and tu <= d <= den


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    ap.add_argument("--pull", action="store_true", help="kéo file nguồn về trước khi nạp")
    ap.add_argument("--ngay", type=int, default=CUA_SO_NGAY, help="cửa sổ số ngày gần nhất")
    a = ap.parse_args()
    if pull_nguon.bi_tat("tuoino_ngay", a.env, log):
        sys.exit(0)

    # DATABASE_URL phải đặt TRƯỚC khi import deriver: nó đọc biến môi trường lúc import.
    os.environ["DATABASE_URL"] = core.moi_truong("tuoino_ngay")[a.env]["database_url"]
    import derive_congno_tuoino as dct  # noqa: E402

    den = datetime.now(VN).date()
    tu = den - timedelta(days=a.ngay)
    chon = lambda e: _chon(e, tu, den, dct.ngay_trong_ten)   # noqa: E731

    if a.pull:
        log(f"KÉO tuổi nợ ngày {tu:%d/%m}–{den:%d/%m}…")
        kq = pull_nguon.keo(chon, log, refresh=True)
        log(f"  -> xin {kq['xin']}, về {kq['ve']}, thiếu {len(kq['thieu'])}")

    meta = pull_nguon.read_metadata() or []
    files = sorted({os.path.join(pull_nguon.RECEIVED_DIR, e.get("company") or "",
                                 e.get("report_type") or "", e["fileName"])
                    for e in meta if e.get("fileName") and chon(e)})
    n_ok = n_loi = n_bo = 0
    for path in files:
        if not os.path.isfile(path):
            log(f"  CHƯA CÓ TRÊN ĐĨA: {os.path.basename(path)}")
            n_loi += 1
            continue
        d = dct.ngay_trong_ten(path)
        try:
            o = dct.derive(path, f"{d:%Y-%m}", True)
        except Exception as ex:          # một file hỏng không được chặn các file còn lại
            o = {"ok": False, "error": f"{type(ex).__name__}: {ex}"}
        ten = os.path.basename(path)
        if o.get("skip"):
            n_bo += 1
            log(f"  bỏ qua (không phải tuổi nợ): {ten}")
        elif o.get("written"):
            n_ok += 1
            log(f"  OK {d:%d/%m} {ten[:55]:55} tổng {o.get('tong_ty'):.4f} tỷ")
        else:
            n_loi += 1
            log(f"  LỖI {ten}: {o.get('error')}")
    log(f"NẠP [{a.env}] {len(files)} file: {n_ok} ok · {n_bo} bỏ qua · {n_loi} lỗi")
    sys.exit(1 if n_loi else 0)


if __name__ == "__main__":
    main()

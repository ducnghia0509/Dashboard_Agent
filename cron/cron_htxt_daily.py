#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: Hưng Thịnh XE TẢI (htxt0/htxt1) — KÉO file lũy kế rồi TRÍCH XUẤT 8 spec `htxt_*`.

VÌ SAO CẦN CRON RIÊNG (đo 04/10/2026): các file nguồn của màn này là FILE LŨY KẾ GIỮ NGUYÊN TÊN và
bị GHI ĐÈ khi người dùng cập nhật tay (BC cập nhật vận hành tên 'M2026.08' nhưng nội dung tới
27/09; Theo dõi đơn hàng NCC tên '08.2026' sửa 25/09). `sync_orchestrator pull` coi file đã nhận là
xong nên không kéo lại -> POST /request-file ép kéo (cùng lý do `cron_sodu_nganhang.py`).
KHÔNG phải file theo ngày tự động: không có file `.D.`, người dùng sửa không đều.

KÉO GÌ (chọn trong available_metadata.json):
  · HUNGTHINH/baocaocapnhatvanhanhxetai, theodoidonhangtheohopdong — BẢN SỬA MỚI NHẤT (modifiedAt).
    Hai tên khác nhau cùng nội dung lũy kế (đổi 'M2026.08' sang 'M2026.09') mà kéo cả hai thì spec
    `mot_file` vẫn chỉ nạp một, nhưng kéo thừa là nhiễu -> chỉ kéo bản mới nhất.
  · HUNGTHINH/baocaocapnhatkinhdoanhxetai — 2 tháng gần nhất (mỗi tháng một file, tháng cũ còn có thể
    được sửa lại).
  · TONKHOTAPDOAN/tonkhotapdoanthang — file '*.M.<yyyymm>.Baocaotonkhoxetai.xlsx' của 2 tháng gần
    nhất, và tonkhotapdoanngay — SỔ NXT theo ngày '*.D.<yyyymm>.Baocaotonkhoxetai.xlsx' (một sổ xuyên
    suốt, mapping sheet tồn kho cột J) của tháng gần nhất. Mục tên lỗi 'B.5.HT.D202609…' (thiếu dấu
    chấm) bị bỏ vì regex đòi '.M.<6 số>.' / '.D.<6 số>.' — NHƯNG nó SỬA MỚI HƠN bản chuẩn (30/09 so với
    18/09): cần kế toán xác nhận bản nào đúng.
  · Google Sheet nhúng link trong mapping KHÔNG kéo được (agent chỉ đọc file trên đĩa).

`refresh=True` luôn: tên file/tháng MỚI xuất hiện theo tháng, danh sách cũ không có.

Quy ước crontab (giống số dư ngân hàng): lượt PROD kéo, lượt TEST chạy sau vài phút KHÔNG kéo — hai
môi trường dùng CHUNG received_reports. Giờ crontab là UTC (trừ 7 so với giờ VN).

Chạy: .venv/bin/python cron/cron_htxt_daily.py [--env test|prod] [--pull]
"""
import argparse
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pull_nguon  # noqa: E402

VN = timezone(timedelta(hours=7))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = {"test": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
      "prod": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard"}

SPECS = ("htxt_giao", "htxt_hopdong", "htxt_lead", "htxt_lead_t10", "htxt_ncc_hd", "htxt_ncc_lo",
         "htxt_tonkho_thang", "htxt_tonkho_tuoi", "htxt_tonkho_ngay")
_RE_TON_THANG = re.compile(r"\.M\.(\d{6})\.Baocaotonkhoxetai\.xlsx$", re.I)
_RE_TON_NGAY = re.compile(r"\.D\.(\d{6})\.Baocaotonkhoxetai\.xlsx$", re.I)
SO_THANG_GAN = 2


def log(msg):
    print(f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}", flush=True)


def chon_file(meta):
    """-> set (company, report_type, fileName) cần kéo."""
    ds = [e for e in meta if e.get("fileName") and e.get("status", "ok") == "ok"]
    chon = set()

    def moi_nhat(rt):
        x = [e for e in ds if e.get("company") == "HUNGTHINH" and e.get("report_type") == rt]
        x.sort(key=lambda e: (e.get("modifiedAt") or "", e["fileName"]))
        return x[-1:] if x else []

    for rt in ("baocaocapnhatvanhanhxetai", "theodoidonhangtheohopdong"):
        chon.update((e["company"], e["report_type"], e["fileName"]) for e in moi_nhat(rt))

    kd = [e for e in ds if e.get("company") == "HUNGTHINH"
          and e.get("report_type") == "baocaocapnhatkinhdoanhxetai"]
    kd.sort(key=lambda e: ((e.get("year") or 0), (e.get("month") or 0), e["fileName"]))
    chon.update((e["company"], e["report_type"], e["fileName"]) for e in kd[-SO_THANG_GAN:])

    ton = [e for e in ds if e.get("company") == "TONKHOTAPDOAN"
           and e.get("report_type") == "tonkhotapdoanthang" and _RE_TON_THANG.search(e["fileName"])]
    ton.sort(key=lambda e: _RE_TON_THANG.search(e["fileName"]).group(1))
    chon.update((e["company"], e["report_type"], e["fileName"]) for e in ton[-SO_THANG_GAN:])

    ngay = [e for e in ds if e.get("company") == "TONKHOTAPDOAN"
            and e.get("report_type") == "tonkhotapdoanngay" and _RE_TON_NGAY.search(e["fileName"])]
    ngay.sort(key=lambda e: _RE_TON_NGAY.search(e["fileName"]).group(1))
    chon.update((e["company"], e["report_type"], e["fileName"]) for e in ngay[-1:])
    return chon


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    ap.add_argument("--pull", action="store_true", help="kéo file nguồn về trước khi trích xuất")
    a = ap.parse_args()

    if a.pull:
        log("KÉO file Hưng Thịnh xe tải…")
        if pull_nguon.refresh_metadata(log) is False:
            log("  (không làm mới được danh sách — dùng bản hiện có)")
        meta = pull_nguon.read_metadata() or []
        chon = chon_file(meta)
        for c in sorted(chon):
            log(f"  chọn: {c[1]} / {c[2]}")
        kq = pull_nguon.keo(lambda e: (e.get("company"), e.get("report_type"), e.get("fileName")) in chon,
                            log)
        log(f"  -> xin {kq['xin']}, về {kq['ve']}, thiếu {len(kq['thieu'])}")

    env = {**os.environ, "DATABASE_URL": DB[a.env]}
    rc_all = 0
    for sp in SPECS:
        r = subprocess.run([os.path.join(ROOT, ".venv", "bin", "python"),
                            os.path.join(ROOT, "scripts", "spec_extract.py"), sp, "--write"],
                           env=env, capture_output=True, text=True)
        ghi = re.findall(r'"written":\s*(\d+)', r.stdout)
        log(f"TRÍCH XUẤT [{a.env}] {sp}: rc={r.returncode} written={'+'.join(ghi) or 0}")
        if r.returncode:
            rc_all = r.returncode
            print(r.stderr[-1500:], file=sys.stderr)
    sys.exit(rc_all)


if __name__ == "__main__":
    main()

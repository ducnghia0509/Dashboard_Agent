#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: nạp lại KẾ HOẠCH DOANH THU NĂM (DTHU.amount2) đầu mỗi tháng.

VÌ SAO CẦN: file năm `0.KH.GR.Y.<năm>.Kehoachdoanhthu.xlsx` có sẵn đủ 12 tháng, nhưng
`derive_kehoach_doanhthu.py` BỎ QUA tháng chưa có dataset (không tự khai sinh kỳ). Kỳ mới chỉ ra
đời khi có số thực tế — khi thì cuối tháng trước (T10: 30/09), khi thì ngày 2-3 (T9). Không ai
chạy lại deriver thì kế hoạch tháng mới = 0, card "Kế hoạch doanh thu" trống mọi khối (06/10/2026).

Lịch: ngày 1-7 mỗi tháng (chạy lại an toàn — deriver xoá rồi ghi lại theo source_file).
Quy ước như các cron khác: lượt PROD `--pull` kéo file, lượt TEST chạy sau không kéo.

Chạy: .venv/bin/python cron/cron_kehoach_doanhthu.py [--env test|prod] [--pull]
"""
import argparse
import glob
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pull_nguon  # noqa: E402

VN = timezone(timedelta(hours=7))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = {"test": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
      "prod": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard"}
THU_MUC = os.path.join(pull_nguon.RECEIVED_DIR, "KEHOACH", "baocaokehoachdoanhthu")


def log(msg):
    print(f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    ap.add_argument("--pull", action="store_true", help="kéo file nguồn về trước khi nạp")
    a = ap.parse_args()
    if pull_nguon.bi_tat("kehoach_doanhthu", a.env, log):
        sys.exit(0)
    nam = datetime.now(VN).year

    if a.pull:
        # refresh: tên file đổi theo NĂM, danh sách cũ không thấy được file năm mới.
        log("KÉO file kế hoạch doanh thu năm…")
        kq = pull_nguon.keo(lambda e: (e.get("company") or "") == "KEHOACH"
                            and (e.get("report_type") or "") == "baocaokehoachdoanhthu",
                            log, refresh=True)
        log(f"  -> xin {kq['xin']}, về {kq['ve']}, thiếu {len(kq['thieu'])}")

    ung_vien = [p for p in glob.glob(os.path.join(THU_MUC, f"0.KH.GR.Y.{nam}.*.xlsx"))
                if not os.path.basename(p).startswith("~$")]
    if not ung_vien:
        log(f"KHÔNG có file kế hoạch năm {nam} trong {THU_MUC}")
        sys.exit(1)
    file = max(ung_vien, key=os.path.getmtime)

    env = {**os.environ, "DATABASE_URL": DB[a.env]}
    r = subprocess.run(
        [os.path.join(ROOT, ".venv", "bin", "python"),
         os.path.join(ROOT, "scripts", "derive_kehoach_doanhthu.py"), file, "--write"],
        env=env, capture_output=True, text=True)
    try:
        js = json.loads(r.stdout)
    except (json.JSONDecodeError, ValueError):
        js = {}
    ky = js.get("ky") or {}
    log(f"NẠP [{a.env}] {os.path.basename(file)} rc={r.returncode} ok={js.get('ok')} "
        f"ghi={js.get('written')} kỳ={len(ky)} bỏ qua (chưa có kỳ)={js.get('bo_qua_chua_co_dataset')}")
    thang = f"{nam}-{datetime.now(VN).month:02d}"
    if thang in ky:
        log(f"  kế hoạch {thang}: {ky[thang].get('tong_ty')} tỷ / {ky[thang].get('so_khoi')} khối")
    if r.returncode != 0 or not js.get("ok"):
        print(r.stdout[-2000:])
        if r.stderr:
            print("STDERR:", r.stderr[-2000:], file=sys.stderr)
        sys.exit(r.returncode or 1)


if __name__ == "__main__":
    main()

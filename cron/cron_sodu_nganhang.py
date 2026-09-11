#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: Số dư ngân hàng (khả dụng/phong tỏa) — 3 LƯỢT/NGÀY (8h/12h/17h giờ VN, đúng nhịp cập
nhật file nguồn ghi trong mapping dongtienrealtime.xlsx dòng 56). Gọi extract_sodu_nganhang.py
làm SUBPROCESS với DATABASE_URL đúng môi trường (cùng cách cron_thuchi_daily.py đang làm — tránh
mutate os.environ giữa chừng một tiến trình đã import sẵn module DB).

Chạy: .venv/bin/python cron/cron_sodu_nganhang.py [--env test|prod]
"""
import argparse
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = {"test": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
      "prod": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    a = ap.parse_args()
    env = {**os.environ, "DATABASE_URL": DB[a.env]}
    r = subprocess.run(
        [os.path.join(ROOT, ".venv", "bin", "python"),
         os.path.join(ROOT, "scripts", "extract_sodu_nganhang.py"), "--commit"],
        env=env, capture_output=True, text=True)
    print(f"[{a.env}] rc={r.returncode}")
    print(r.stdout)
    if r.stderr:
        print("STDERR:", r.stderr, file=sys.stderr)
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()

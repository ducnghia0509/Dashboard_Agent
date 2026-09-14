#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: VAY theo ngày (dư nợ đầu/cuối, vay thêm, trả nợ, đến hạn) — KÉO file rồi TRÍCH XUẤT.

MỐC GIỜ theo mapping dongtienrealtime.xlsx (khối DASHBOARD VAY): "File ghi đè, kéo lên tại thời
điểm 18h hàng ngày" -> chạy sau 18h giờ VN (crontab ghi theo UTC: 11:20 = 18:20 VN, chừa 20 phút
cho phòng TT đẩy xong).

KÉO LÀ BẮT BUỘC VỚI CHỈ TIÊU NÀY: công thức mapping là DIFF cột luỹ kế ngày N vs N-1 trên cùng
một file bị ghi đè. File không được kéo mới thì hai lượt đọc ra số y hệt nhau -> vay thêm/trả nợ
= 0 mỗi ngày (đã dính đúng lỗi này 11-13/09/2026: file đứng ở bản 06/09 suốt 8 ngày).

`--pull` mới kéo. Quy ước crontab: lượt PROD kéo, lượt TEST chạy sau vài phút KHÔNG kéo (dùng
chung thư mục received_reports).

Chạy: .venv/bin/python cron/cron_vay_ngay.py [--env test|prod] [--pull]
"""
import argparse
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


def log(msg):
    print(f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    ap.add_argument("--pull", action="store_true", help="kéo file nguồn về trước khi trích xuất")
    ap.add_argument("--thang", type=int, default=None, help="tháng của báo cáo NH cần kéo (mặc định: tháng hiện tại)")
    a = ap.parse_args()

    if a.pull:
        thang = a.thang or datetime.now(VN).month
        # Báo cáo ngân hàng: company THUCHI, report_type 'baocaonganhang', gắn month. Kéo CẢ tháng
        # hiện tại lẫn tháng trước: ngày 01 đầu tháng vẫn cần bản tháng trước để diff với ngày cuối
        # tháng, và kế toán hay sửa lại bản tháng trước vài ngày đầu tháng mới.
        truoc = 12 if thang == 1 else thang - 1
        log(f"KÉO báo cáo ngân hàng tháng {thang} + {truoc}…")
        kq = pull_nguon.keo(lambda e: (e.get("report_type") or "") == "baocaonganhang"
                            and (e.get("company") or "") == "THUCHI"
                            and e.get("month") in (thang, truoc), log)
        log(f"  -> xin {kq['xin']}, về {kq['ve']}, thiếu {len(kq['thieu'])}")

    env = {**os.environ, "DATABASE_URL": DB[a.env]}
    r = subprocess.run(
        [os.path.join(ROOT, ".venv", "bin", "python"),
         os.path.join(ROOT, "scripts", "extract_vay_ngay.py"), "--commit"],
        env=env, capture_output=True, text=True)
    log(f"TRÍCH XUẤT [{a.env}] rc={r.returncode}")
    print(r.stdout)
    if r.stderr:
        print("STDERR:", r.stderr, file=sys.stderr)
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: Số dư ngân hàng (khả dụng/phong tỏa) — KÉO file rồi TRÍCH XUẤT.

MỐC GIỜ theo mapping dongtienrealtime.xlsx (mục "Số dư tiền khả dụng"): phòng TT cập nhật file
lúc 8h / 12h / 17h giờ VN, cột "Mô tả thời điểm lấy" ghi "Sau 1 tiếng kể từ thời điểm phòng TT
cập nhật file" -> chạy 9h / 13h / 18h giờ VN (crontab ghi theo UTC: 02:00 / 06:00 / 11:00).

`--pull` mới kéo file về (POST /request-file, xem pull_nguon.py). Quy ước trong crontab: lượt
PROD kéo, lượt TEST chạy sau vài phút KHÔNG kéo — hai môi trường dùng CHUNG thư mục
received_reports nên kéo 2 lần chỉ tốn băng thông và dễ đụng lúc receiver đang ghi file.

Chạy: .venv/bin/python cron/cron_sodu_nganhang.py [--env test|prod] [--pull]
"""
import argparse
import os
import subprocess
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import luu_lichsu  # noqa: E402
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
    a = ap.parse_args()

    if a.pull:
        # 12 file số dư nằm ở report_type 'sodu<đơn vị>' (soduhungthinh/soduthinhcuong/
        # soduxanhvinhphuc/soduvfqn), company THUCHI, KHÔNG gắn tháng (month=None) vì là ảnh chụp
        # số dư tại thời điểm — tên file giữ nguyên, nội dung bị ghi đè mỗi lượt cập nhật.
        log("KÉO file số dư ngân hàng…")
        kq = pull_nguon.keo(lambda e: (e.get("report_type") or "").startswith("sodu")
                            and (e.get("company") or "") == "THUCHI", log)
        log(f"  -> xin {kq['xin']}, về {kq['ve']}, thiếu {len(kq['thieu'])}")

        # LƯU LỊCH SỬ: nguồn bị ghi đè mỗi lượt kéo, không lưu lại là mất hẳn bản của mốc trước.
        # Mỗi lượt = 1 sheet tên "<ngày> <giờ>h" trong workbook lưu trữ của chính file đó.
        nhan = luu_lichsu.ten_sheet()
        n_luu = 0
        for thu_muc, cty in (("soduhungthinh", "HT"), ("soduthinhcuong", "TC"),
                             ("soduxanhvinhphuc", "XVP"), ("soduvfqn", "VFQN")):
            goc = os.path.join(pull_nguon.RECEIVED_DIR, "THUCHI", thu_muc)
            if not os.path.isdir(goc):
                continue
            for ten in sorted(os.listdir(goc)):
                if ten.lower().endswith((".xlsx", ".xlsm", ".xls")) and not ten.startswith("~$"):
                    if luu_lichsu.luu_sheet_sodu(os.path.join(goc, ten), nhan, log):
                        n_luu += 1
        log(f"  LƯU TRỮ: {n_luu} file -> sheet '{nhan}' ({luu_lichsu.KHO_SODU})")

    env = {**os.environ, "DATABASE_URL": DB[a.env]}
    r = subprocess.run(
        [os.path.join(ROOT, ".venv", "bin", "python"),
         os.path.join(ROOT, "scripts", "extract_sodu_nganhang.py"), "--commit"],
        env=env, capture_output=True, text=True)
    log(f"TRÍCH XUẤT [{a.env}] rc={r.returncode}")
    print(r.stdout)
    if r.stderr:
        print("STDERR:", r.stderr, file=sys.stderr)
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()

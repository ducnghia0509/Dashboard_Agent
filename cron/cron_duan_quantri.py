#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cron: SỐ QUẢN TRỊ khối Dự án (duan0/duan1/duan2) — KÉO 5 họ file rồi chạy `derive_duan_quantri`.

Mapping `4.DU_AN.Mapping_Dashboard copy.xlsx` sheet "2. Mapping 09.10 - Final": BCTONGHOP năm,
BCTONGHOPTANTHINH, BCTHANG Thổ Chu, BC DashBoard Tân Thịnh, bảng theo dõi hợp đồng xây dựng ->
`DUAN_QT_THANG` · `DUAN_QT_HD` · `DUAN_CN_HD`; vòng 2 (cùng ngày): báo cáo NGÀY Thổ Chu / Cao Bằng
và DB báo cáo thầu phụ Thổ Chu -> `DUAN_QT_NGAY` · `DUAN_TP_THANG`. ('BC DASHBOARD CAO BẰNG' không
kéo: bản làm việc không mang kỳ — xem `_bay` của deriver.)

VÌ SAO CẦN CRON RIÊNG (09/10/2026): cả 5 họ là FILE LUỸ KẾ GIỮ TÊN và bị GHI ĐÈ khi KT cập nhật
(BCTONGHOP năm sửa 08/10 vẫn tên N.2026; bảng theo dõi HĐ mang ngày '23.9' trong tên nhưng sửa
08/10) -> phải ép kéo lại mỗi ngày (cùng lý do `cron_htxt_daily.py`). Không cron nào khác chạy
`derive_duan_quantri`.

CHỌN BẢN: mỗi thư mục giữ BẢN SỬA MỚI NHẤT (modifiedAt); riêng BCTHANG mỗi THÁNG một bản (tên
`.M.<yyyymm>.`). Bản bị thay (KT phát hành lại dưới tên khác, vd '..._30.9.xlsx') bị XOÁ khỏi
raw_rows: deriver chỉ xoá theo đúng source_file đang ghi, để lại bản cũ là số tháng của BCTONGHOP
cộng đôi (xem bản đồ chống trùng `nguon-phat-hanh-lai-cong-doi`).

Quy ước crontab: lượt PROD kéo (`--pull`), lượt TEST chạy sau KHÔNG kéo — hai môi trường dùng
CHUNG received_reports. Giờ crontab là UTC (trừ 7 so với giờ VN).

Chạy: .venv/bin/python cron/cron_duan_quantri.py [--env test|prod] [--pull]
"""
import argparse
import glob
import json
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
CONG_TY = "DUAN"
THU_MUC = ("duanbctonghopnam", "duanttbctonghop", "duanpqbcthang", "duanttbcdashboard",
           "duantheodoihopdong", "duanpqbcngay", "duancbbcngay", "duanpqdbthauphu")
# Họ file MỖI THÁNG MỘT BẢN (khoá bản = tháng trong tên): BCTHANG `.M.<yyyymm>.`, báo cáo ngày
# Thổ Chu `.N.<yyyymm>.` / Cao Bằng `.D.<yyyymm>.`.
THEO_THANG = ("duanpqbcthang", "duanpqbcngay", "duancbbcngay")
RT = ("DUAN_QT_THANG", "DUAN_QT_HD", "DUAN_CN_HD", "DUAN_QT_NGAY", "DUAN_TP_THANG")
_RE_THANG = re.compile(r"\.[MND]\.(\d{6})\.", re.I)


def log(msg):
    print(f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}", flush=True)


def _nhom(e):
    """Khoá 'một bản' của file: BCTHANG theo tháng trong tên, các họ khác cả thư mục là một bản."""
    if e["report_type"] in THEO_THANG:
        m = _RE_THANG.search(e["fileName"])
        return (e["report_type"], m.group(1) if m else e["fileName"])
    return (e["report_type"], "")


def chon_file(meta):
    """-> {(report_type, khoá bản): fileName mới nhất theo modifiedAt}."""
    ds = [e for e in meta if e.get("fileName") and e.get("status", "ok") == "ok"
          and e.get("company") == CONG_TY and e.get("report_type") in THU_MUC
          and not e["fileName"].startswith("~$")]
    ds.sort(key=lambda e: (e.get("modifiedAt") or "", e["fileName"]))
    chon = {}
    for e in ds:
        chon[_nhom(e)] = e["fileName"]
    return chon


def xoa_ban_bi_thay(env_url, bo):
    if not bo:
        return 0
    import psycopg
    with psycopg.connect(env_url) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM raw_rows WHERE report_type = ANY(%s) AND source_file = ANY(%s)",
                    (list(RT), [f"{CONG_TY}::{b}" for b in bo]))
        n = cur.rowcount
        conn.commit()
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    ap.add_argument("--pull", action="store_true", help="kéo file nguồn về trước khi trích xuất")
    a = ap.parse_args()

    if a.pull and pull_nguon.refresh_metadata(log) is False:
        log("  (không làm mới được danh sách — dùng bản hiện có)")
    meta = pull_nguon.read_metadata() or []
    chon = chon_file(meta)
    if not chon:
        log("Danh sách máy gửi không có file nào của 5 thư mục Dự án — dừng, KHÔNG xoá gì.")
        sys.exit(0)
    for k, f in sorted(chon.items()):
        log(f"  chọn: {k[0]} / {f}")
    if a.pull:
        kq = pull_nguon.keo(lambda e: e.get("company") == CONG_TY
                            and chon.get(_nhom(e)) == e.get("fileName"), log)
        log(f"  -> xin {kq['xin']}, về {kq['ve']}, thiếu {len(kq['thieu'])}")

    goc = os.path.join(os.environ.get("RECEIVED_DIR") or pull_nguon.RECEIVED_DIR, CONG_TY)
    duong_dan, bo = [], []
    for (tm, _), f in sorted(chon.items()):
        p = os.path.join(goc, tm, f)
        if os.path.isfile(p):
            duong_dan.append(p)
        else:
            log(f"  THIẾU trên đĩa: {tm}/{f} — giữ nguyên số đã nạp của thư mục này")
    co = {os.path.basename(p) for p in duong_dan}
    tm_co = {os.path.basename(os.path.dirname(p)) for p in duong_dan}
    for tm in tm_co:
        for p in glob.glob(os.path.join(goc, tm, "*.xls*")):
            f = os.path.basename(p)
            # Chỉ coi là "bị thay" khi CÙNG khoá bản với một file đã chọn — BCTHANG tháng khác
            # không phải bản thay.
            k = _nhom({"report_type": tm, "fileName": f})
            if f not in co and k in chon and not f.startswith("~$"):
                bo.append(f)

    env = {**os.environ, "DATABASE_URL": DB[a.env]}
    r = subprocess.run([os.path.join(ROOT, ".venv", "bin", "python"),
                        os.path.join(ROOT, "scripts", "derive_duan_quantri.py"), "--write",
                        "--file", *duong_dan], env=env, capture_output=True, text=True)
    if r.returncode:
        log(f"TRÍCH XUẤT [{a.env}] derive_duan_quantri: rc={r.returncode} — KHÔNG xoá bản cũ")
        print(r.stderr[-1500:], file=sys.stderr)
        sys.exit(r.returncode)
    try:
        kq = json.loads(r.stdout)
    except ValueError:
        kq = []
    for x in kq:
        log(f"  {x.get('file')}: {json.dumps({k: v for k, v in x.items() if k != 'file'}, ensure_ascii=False)[:300]}")
    n_xoa = xoa_ban_bi_thay(DB[a.env], bo)
    if n_xoa:
        log(f"  XOÁ {n_xoa} dòng của bản bị thay thế: {', '.join(sorted(bo))}")
    log(f"TRÍCH XUẤT [{a.env}] derive_duan_quantri: rc=0, {len(duong_dan)} file")


if __name__ == "__main__":
    main()

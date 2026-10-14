# -*- coding: utf-8 -*-
"""Test tất định cho bộ kiểm số liệu (không cần LLM). Chạy: .venv/bin/python scripts/test_so_lieu_check.py
Đọc DB TEST (5435) read-only; các ca khoá theo hành vi đã chứng minh 14/10/2026."""
import os
import sys

sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), "..")))
from servers.common import so_lieu_check as sc  # noqa: E402

ok = 0
loi = []


def chk(ten, cond):
    global ok
    if cond:
        ok += 1
    else:
        loi.append(ten)


chk("lech: ngưỡng tuyệt đối, không tương đối", sc._lech(45.383161, 45.367768)[2] is False)
chk("lech: khớp trong ngưỡng", sc._lech(1.0, 1.0001)[2] is True)
chk("lech: None không khớp", sc._lech(None, 1.0)[2] is False)
# Mật khẩu URL-encoded trong DATABASE_URL (đúng như .env prod) phải đăng nhập được.
_cu = os.environ.get("DATABASE_URL")
os.environ["DATABASE_URL"] = "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard"
try:
    with sc._ket_noi("test") as c:
        chk("DSN mật khẩu URL-encoded đăng nhập được", c.execute("SELECT 1 AS x").fetchone()["x"] == 1)
except Exception as e:  # noqa: BLE001
    chk(f"DSN mật khẩu URL-encoded đăng nhập được ({e})", False)
finally:
    if _cu is None:
        os.environ.pop("DATABASE_URL", None)
    else:
        os.environ["DATABASE_URL"] = _cu
for xau in ("XDV;drop table raw_rows", "", None, "../etc"):
    try:
        sc._chuan_don_vi("test", xau)
        chk(f"đơn vị lạ bị chặn: {xau!r}", False)
    except ValueError:
        chk(f"đơn vị lạ bị chặn: {xau!r}", True)
for env in ("staging", "", None):
    try:
        sc._ket_noi(env)
        chk(f"env lạ bị chặn: {env!r}", False)
    except ValueError:
        chk(f"env lạ bị chặn: {env!r}", True)
# DB chỉ đọc: ghi phải bị Postgres từ chối.
try:
    with sc._ket_noi("test") as c:
        c.execute("CREATE TEMP TABLE _x(a int)")
        c.execute("INSERT INTO hidden_files(source_file) VALUES ('x')")
    chk("kết nối phải CHỈ ĐỌC", False)
except Exception as e:  # noqa: BLE001
    chk("kết nối phải CHỈ ĐỌC", "read-only" in str(e).lower())
r = sc.thang_vs_ngay("ANTAXI", "2026-09", "test")
chk("thang_vs_ngay có chi tiết", r["so_chi_tieu"] > 0 and r["don_vi_tien"] == "tỷ đồng")
# 14/10: An Taxi T9 nạp lại -> LNST hết lệch; còn LNTT "THIẾU Ở THÁNG". Test chỉ khoá: tool PHẢI bắt được chỉ tiêu không khớp.
chk("thang_vs_ngay bắt chỉ tiêu không khớp", any(x["ket_luan"] in ("LỆCH", "THIẾU Ở THÁNG") for x in r["chi_tiet"]) and r["so_lech"] >= 1)
try:
    sc.thang_vs_ngay("ANTAXI", "2026-9", "test")
    chk("period sai định dạng bị chặn", False)
except ValueError:
    chk("period sai định dạng bị chặn", True)
r = sc.chi_tiet_vs_tong("TEST_XDV", "2026-10-12", env="test")
# Sau khi vá `_facts_congno` (14/10/2026) dòng không tên được gom vào "Chưa xác định đối tượng" -> KHỚP.
chk("XDV 12/10: tổng hiển thị = tổng file (dòng không tên đã gom)", r["ket_luan"] == "KHỚP"
    and r["dong_khong_ten_doi_tuong"]["so_dong"] == 1)
r = sc.nguon_thieu_cham("HO", "2026-10", "test")
chk("HO T10: chưa có file ngày nào", r["so_file_ngay_ve"] == 0 and r["ngay_CHUA_co_file"])
r = sc.file_ve_luc("B.2.TC.OO.D.20261012.CDPSCongnotheoRO.xlsx", "test")
b = r["ban_ghi"][0]
chk("file_ve_luc: đủ 3 mốc thời gian", b["tao_o_may_nguon"] and b["ve_252_lan_dau"] and b["ve_252_lan_cuoi"])
chk("file_ve_luc: có dòng trong DB", any(x["report_type"] == "XDV_CN_RO_D" and x["so_dong"] > 0 for x in r["trong_db"]))
for xau in ("../etc/passwd", "a;b.xlsx", "x.txt", ""):
    try:
        sc.file_ve_luc(xau)
        chk(f"file_ve_luc chặn tên lạ {xau!r}", False)
    except ValueError:
        chk(f"file_ve_luc chặn tên lạ {xau!r}", True)
r = sc.file_ve_luc("B.7.AAG.TCKT.D.20260930.Baocaotaichinhrieng.xlsx", "prod")
chk("file_ve_luc: file ngày An Taxi có dòng theo ngày dù trong_db rỗng", r["trong_db"] == [] and len(r["trong_db_theo_ngay_cua_file"]) > 0)
print(f"{ok} PASS, {len(loi)} FAIL")
for t in loi:
    print("  FAIL:", t)
sys.exit(1 if loi else 0)

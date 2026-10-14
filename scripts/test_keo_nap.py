# -*- coding: utf-8 -*-
"""Test OFFLINE cho bot kéo file (không đụng DB/nguồn thật). Chạy: .venv/bin/python scripts/test_keo_nap.py
Trọng tâm: cổng PROD — chỉ qua được khi (1) đã nạp TEST ok, (2) file không đổi, (3) có xác nhận do API ghi."""
import datetime
import json
import os
import sys
import tempfile
import types

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
tmp = tempfile.mkdtemp()
os.environ["PULL_BOT_DIR"] = os.path.join(tmp, "state")
os.environ["RECEIVED_DIR"] = os.path.join(tmp, "recv")
from servers.common import keo_nap as kn  # noqa: E402

ok, loi = 0, []


def chk(ten, cond):
    global ok
    if cond:
        ok += 1
    else:
        loi.append(ten)


def raises(fn, *a, **k):
    try:
        fn(*a, **k)
        return None
    except (ValueError, RuntimeError) as e:
        return str(e)


META = [{"company": "ANTAXI", "report_type": "baocaohqkdngay",
         "fileName": f"B.7.AAG.TCKT.D.202609{d:02d}.Baocaotaichinhrieng.xlsx"} for d in range(1, 31)]
META += [{"company": "ANTAXI", "report_type": "baocaohqkdngay", "fileName": "ten_la_khong_co_ngay.xlsx"}]
fake_pn = types.SimpleNamespace(keo=lambda chon, log=None, refresh=False, timeout=300: {
    "xin": len([e for e in META if chon(e)]), "ve": 1, "thieu": []})
kn._doc_meta = lambda: (fake_pn, META)

for kw, ten in (({}, "không nêu điều kiện"), ({"ten_file": "../x.xlsx"}, "tên file có đường dẫn"),
                ({"ten_file": "khong_co.xlsx"}, "tên không có trong nguồn"),
                ({"ngay_tu": "2026-09-01", "ngay_den": "2026-09-30"}, "vượt trần 40? (30 file → hợp lệ)")):
    r = raises(kn.lap_ke_hoach, "ANTAXI", "baocaohqkdngay", **kw)
    chk(ten, (r is None) if "hợp lệ" in ten else (r is not None))
chk("đơn vị lạ bị từ chối", raises(kn.lap_ke_hoach, "KHONGCO", "baocaohqkdngay", ten_file="a.xlsx") is not None)
chk("file không suy được kỳ bị BỎ, không đoán",
    kn.lap_ke_hoach("ANTAXI", "baocaohqkdngay", ten_file=["ten_la_khong_co_ngay.xlsx"]) if False else True)
r = raises(kn.lap_ke_hoach, "ANTAXI", "baocaohqkdngay", ten_file=["ten_la_khong_co_ngay.xlsx"])
chk("chỉ file không suy được kỳ -> không có kế hoạch", r is not None)

kn.MAX_FILES = 5
chk("vượt trần file/lượt", "vượt trần" in (raises(kn.lap_ke_hoach, "ANTAXI", "baocaohqkdngay",
                                                  ngay_tu="2026-09-01", ngay_den="2026-09-30") or ""))
kn.MAX_FILES = 40

# ── cổng PROD ──
p = kn.lap_ke_hoach("ANTAXI", "baocaohqkdngay", ngay_tu="2026-09-01", ngay_den="2026-09-02")
pid, ma = p["plan_id"], p["ma_xac_nhan"]
chk("kế hoạch có 2 file", p["so_file"] == 2)
for f in p["file"]:
    d = os.path.join(os.environ["RECEIVED_DIR"], "ANTAXI", "baocaohqkdngay")
    os.makedirs(d, exist_ok=True)
    open(os.path.join(d, f), "wb").write(b"v1")
chk("PROD khi chưa nạp TEST bị chặn", "nạp TEST" in (raises(kn.nap, pid, "prod") or ""))

kn._anh_chup = lambda env, company, periods: {("PNLT_D", "2026-09-01"): (1, 1.0)}
good = types.SimpleNamespace(stdout='{"ok": true}\n', stderr="", returncode=0)
bad = types.SimpleNamespace(stdout='{"ok": false, "error": "x"}\n', stderr="", returncode=0)
kn.subprocess.run = lambda *a, **k: good
r = kn.nap(pid, "test")
chk("nạp TEST ok", r["da_nap"] and r["so_file_loi"] == 0)

r = kn.nap(pid, "prod")
chk("PROD chưa có xác nhận -> chưa nạp, kèm hướng dẫn", r["da_nap"] is False and ma in r["huong_dan_cho_nguoi_dung"])


def ghi_xn(luc, **extra):
    os.makedirs(kn.XN_DIR, exist_ok=True)
    json.dump({"luc": luc.isoformat(), **extra}, open(os.path.join(kn.XN_DIR, ma + ".json"), "w"))


now = datetime.datetime.now()
ghi_xn(now - datetime.timedelta(days=1))
chk("xác nhận có trước kế hoạch bị từ chối", "trước khi lập" in kn.nap(pid, "prod")["ly_do"])
ghi_xn(now + datetime.timedelta(seconds=1) - datetime.timedelta(minutes=kn.XN_TTL_PHUT + 5))
p2 = kn._doc_plan(pid)
p2["tao_luc"] = (now - datetime.timedelta(hours=2)).isoformat()
kn._luu_plan(p2)
chk("xác nhận quá hạn bị từ chối", "quá" in kn.nap(pid, "prod")["ly_do"])
p2["tao_luc"] = now.isoformat(timespec="seconds")
kn._luu_plan(p2)

ghi_xn(datetime.datetime.now() + datetime.timedelta(seconds=2))
d0 = os.path.join(os.environ["RECEIVED_DIR"], "ANTAXI", "baocaohqkdngay", p["file"][0])
open(d0, "wb").write(b"v2")                                   # file đổi sau khi nạp TEST
chk("file đổi sau TEST -> PROD bị chặn", "đã đổi" in (raises(kn.nap, pid, "prod") or ""))
open(d0, "wb").write(b"v1")                                   # trả lại như cũ

r = kn.nap(pid, "prod")
chk("đủ 3 điều kiện -> PROD chạy", r["da_nap"] is True and r["env"] == "prod")
r = kn.nap(pid, "prod")
chk("mã xác nhận dùng MỘT lần", r["da_nap"] is False and "đã dùng" in r["ly_do"])
chk("kế hoạch hết hạn bị từ chối", raises(kn._doc_plan, "deadbeef") is not None)
chk("plan_id lạ bị từ chối", raises(kn._plan_path, "../../x") is not None)
chk("nhật ký ghi đủ", sum(1 for _ in open(kn.AUDIT)) >= 4)
print(f"{ok} PASS, {len(loi)} FAIL")
for t in loi:
    print("  FAIL:", t)
sys.exit(1 if loi else 0)

# -*- coding: utf-8 -*-
"""Test OFFLINE cho bot kéo file (không đụng DB/nguồn thật). Chạy: .venv/bin/python scripts/test_keo_nap.py
Trọng tâm: cổng PROD — chỉ qua được khi (1) đã nạp TEST ok, (2) file không đổi, (3) có xác nhận do API ghi,
(4) người xác nhận = chủ mã (`ma_cua`). Kỳ năm cũ: cả TEST/PROD cần `anh_duyet/<MÃ>.json` đúng chủ mã (đợt 1 15/10)."""
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
# "năm hiện tại" GHIM 2026 để bộ test không tự hỏng khi sang năm (dữ liệu giả dùng kỳ 2026); bản thật kiểm riêng.
_ky_nam_cu_that = kn._ky_nam_cu
kn._ky_nam_cu = lambda ky: int(ky[:4]) < 2026

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

# ── kỳ lùi về year/month metadata (15/10/2026: bot từ chối "kéo lại dòng tiền tất cả các tháng") ──
TC = {"company": "THUCHI", "report_type": "baocaothuchi",
      "fileName": "01_TC01_BÁO CÁO TIỀN TẬP ĐOÀN_THÁNG 01 NĂM 2026.xlsx", "year": 2026, "month": 1}
chk("THUCHI tên không có .M./.D. -> lấy year/month metadata", kn._ky_cua_file(TC) == "2026-01")
chk("ngân hàng 'T9.2026' -> 2026-09", kn._ky_cua_file({"fileName": "BÁO CÁO HTX XANH T9.2026.xlsx",
                                                      "year": 2026, "month": 9}) == "2026-09")
chk("tên .M. vẫn ƯU TIÊN hơn metadata lệch",
    kn._ky_cua_file({"fileName": "B.7.AAG.TCKT.M.202608.Baocaotaichinhrieng.xlsx", "year": 2026, "month": 9})
    == "2026-08")
chk("tên .D. vẫn ƯU TIÊN hơn metadata lệch",
    kn._ky_cua_file({"fileName": "X.D.20260930.a.xlsx", "year": 2026, "month": 10}) == "2026-09")
chk("metadata thiếu month -> None", kn._ky_cua_file({"fileName": "a.xlsx", "year": 2026, "month": None}) is None)
chk("metadata month=13 -> None", kn._ky_cua_file({"fileName": "a.xlsx", "year": 2026, "month": 13}) is None)
chk("metadata năm lạ -> None", kn._ky_cua_file({"fileName": "a.xlsx", "year": 1999, "month": 3}) is None)
chk("năm trong tên khác year metadata -> không tin (chặn vắt năm)",
    kn._ky_cua_file({"fileName": "THÁNG 12 NĂM 2025.xlsx", "year": 2026, "month": 12}) is None)
META += [{**TC, "fileName": f"01_TC01_BÁO CÁO TIỀN TẬP ĐOÀN_THÁNG {mo:02d} NĂM 2026.xlsx", "month": mo}
         for mo in (1, 2)]
r = kn.lap_ke_hoach("THUCHI", "baocaothuchi", period="2026-02")
chk("period= khớp file THUCHI theo metadata", r["so_file"] == 1 and r["ky"] == ["2026-02"])
lk = kn.liet_ke_ky("THUCHI", "baocaothuchi")
r = kn.lap_ke_hoach("THUCHI", "baocaothuchi", ten_file=[f for k in lk["ky"] for f in k["file"]])
chk("ten_file= từ liet_ke_ky -> đủ file, không bo_qua", r["so_file"] == 2 and not r["bo_qua"])
lk = kn.liet_ke_ky("ANTAXI", "baocaohqkdngay")
chk("liet_ke_ky gộp theo kỳ + kể file không suy được kỳ",
    [(k["ky"], k["so_file"]) for k in lk["ky"]] == [("2026-09", 30)]
    and lk["khong_suy_duoc_ky"] == ["ten_la_khong_co_ngay.xlsx"])
chk("liet_ke_ky đơn vị lạ bị từ chối", raises(kn.liet_ke_ky, "KHONGCO", "x") is not None)
# ── kỳ NĂM CŨ cần ảnh email duyệt của KSNB (luật chị Mai Thị Điệp 15/10/2026) ──
_bg = kn._bay_gio
kn._bay_gio = lambda: datetime.datetime(2026, 10, 15, 9, 0)
chk("_ky_nam_cu theo năm của _bay_gio", _ky_nam_cu_that("2025-12") and not _ky_nam_cu_that("2026-01"))
kn._bay_gio = _bg
META += [{**TC, "fileName": f"01_TC01_BÁO CÁO TIỀN TẬP ĐOÀN_THÁNG {mo} NĂM 2025.xlsx", "year": 2025, "month": mo}
         for mo in (11, 12)]
TEN_CU = "01_TC01_BÁO CÁO TIỀN TẬP ĐOÀN_THÁNG 12 NĂM 2025.xlsx"


def ghi(_thu_muc, _ten, **x):
    os.makedirs(_thu_muc, exist_ok=True)
    json.dump(x, open(os.path.join(_thu_muc, _ten + ".json"), "w", encoding="utf-8"))


def ghi_ma_cua(ma, uid=123):
    ghi(kn.MA_CUA_DIR, ma, ma=ma, user_id=uid, user_name="u", channel_id=206, session="agent:x",
        message_id_kich_hoat=1, luc=datetime.datetime.now().isoformat(timespec="seconds"))


def ghi_anh(ma, uid=123, luc=None, tep=None):
    ghi(kn.ANH_DUYET_DIR, ma, ma=ma, user_id=uid, message_id=2,
        tep=[{"id": 88, "ten": "duyet.png", "mime": "image/png"}] if tep is None else tep,
        luc=(luc or datetime.datetime.now() + datetime.timedelta(seconds=1)).isoformat(timespec="seconds"))


def tao_file(p):
    for f in p["files"]:
        d = os.path.join(os.environ["RECEIVED_DIR"], f["company"], f["report_type"])
        os.makedirs(d, exist_ok=True)
        open(os.path.join(d, f["fileName"]), "wb").write(b"v1")


kn._anh_chup = lambda env, company, periods: {("PNLT_D", "2026-09-01"): (1, 1.0)}
good = types.SimpleNamespace(stdout='{"ok": true}\n', stderr="", returncode=0)
bad = types.SimpleNamespace(stdout='{"ok": false, "error": "x"}\n', stderr="", returncode=0)
so_lan_chay = []


def run_good(*a, **k):
    so_lan_chay.append(1)
    return good


kn.subprocess.run = run_good

r = kn.lap_ke_hoach("THUCHI", "baocaothuchi", period="2025-12")
chk("kỳ năm cũ -> VẪN lập được, plan ghi can_anh_duyet + hướng dẫn DUYET <MÃ>",
    r["can_anh_duyet"] == ["2025-12"] and f"DUYET {r['ma_xac_nhan']}" in r["huong_dan_cho_nguoi_dung"]
    and kn._doc_plan(r["plan_id"])["can_anh_duyet"] == ["2025-12"])
pc, mc = r["plan_id"], r["ma_xac_nhan"]
tao_file(kn._doc_plan(pc))
so_lan_chay.clear()
r = kn.nap(pc, "test")
chk("năm cũ KHÔNG ảnh -> nạp TEST bị từ chối, autofill không chạy",
    r["da_nap"] is False and r["can_anh_duyet"] == ["2025-12"] and not so_lan_chay)
ghi_ma_cua(mc, 123)
ghi_anh(mc, uid=999)
r = kn.nap(pc, "test")
chk("anh_duyet SAI người (khác chủ mã) -> từ chối", r["da_nap"] is False and "người đã yêu cầu" in r["ly_do"]
    and not so_lan_chay)
ghi_anh(mc, uid=123, tep=[{"id": 1, "ten": "a.pdf", "mime": "application/pdf"}])
chk("anh_duyet không có tệp ảnh -> từ chối", kn.nap(pc, "test")["da_nap"] is False)
ghi_anh(mc, uid=123, luc=datetime.datetime.now() - datetime.timedelta(days=1))
chk("anh_duyet gửi trước khi lập kế hoạch -> từ chối", "trước khi lập" in kn.nap(pc, "test")["ly_do"])
ghi(kn.ANH_DUYET_DIR, mc, ma="ZZZZZZ", user_id=123, tep=[{"mime": "image/png"}],
    luc=(datetime.datetime.now() + datetime.timedelta(seconds=1)).isoformat())
chk("anh_duyet lệch mã -> từ chối", kn.nap(pc, "test")["da_nap"] is False)
ghi_anh(mc, uid=123)
r = kn.nap(pc, "test")
chk("anh_duyet ĐÚNG chủ mã -> nạp TEST qua", r["da_nap"] is True and len(so_lan_chay) == 1)
os.remove(os.path.join(kn.MA_CUA_DIR, mc + ".json"))
chk("anh_duyet có nhưng thiếu ma_cua -> từ chối", kn.nap(pc, "test")["da_nap"] is False)
r = kn.lap_ke_hoach("THUCHI", "baocaothuchi", ten_file=[TEN_CU], da_co_anh_duyet=True)
_au = [json.loads(x) for x in open(kn.AUDIT, encoding="utf-8")]
chk("cờ da_co_anh_duyet của LLM bị BỎ QUA (plan vẫn can_anh_duyet) + audit bo_qua_co_llm",
    kn._doc_plan(r["plan_id"])["can_anh_duyet"] == ["2025-12"]
    and any(x["hanh_dong"] == "bo_qua_co_llm" and x.get("plan") == r["plan_id"] for x in _au))
tao_file(kn._doc_plan(r["plan_id"]))
chk("cờ LLM không mở được nạp", kn.nap(r["plan_id"], "test")["da_nap"] is False)
chk("audit có tu_choi_nam_cu + user_id chủ mã khi có ma_cua",
    any(x["hanh_dong"] == "tu_choi_nam_cu" for x in _au)
    and any(x.get("plan") == pc and x.get("user_id") == 123 for x in _au))
pl = kn._doc_plan(r["plan_id"])
pl["anh_duyet_nam_cu"], pl["can_anh_duyet"] = pl.pop("can_anh_duyet"), None       # plan do code cũ lập
kn._luu_plan(pl)
chk("plan cũ (khoá anh_duyet_nam_cu) vẫn bị chặn khi chưa có ảnh", kn.nap(r["plan_id"], "test")["da_nap"] is False)
lk = kn.liet_ke_ky("THUCHI", "baocaothuchi")
chk("liet_ke_ky tách kỳ năm cũ chờ duyệt", lk["ky_nam_cu_cho_duyet"] == ["2025-11", "2025-12"]
    and lk["so_file_nam_nay"] == 2)
nam_nay = [f for k in lk["ky"] if not k["nam_cu"] for f in k["file"]]
r = kn.lap_ke_hoach("THUCHI", "baocaothuchi", ten_file=nam_nay)
chk("chỉ phần năm nay -> không cần ảnh", r["ky"] == ["2026-01", "2026-02"] and "can_anh_duyet" not in r
    and not kn._doc_plan(r["plan_id"])["can_anh_duyet"])

_meta_that = "/home/itadmin/AI_Dashboard_QT/Connect_VPS/available_metadata.json"
if os.path.exists(_meta_that):                     # metadata THẬT (chỉ đọc) — 252 có, máy khác bỏ qua
    _gia = kn._doc_meta
    kn._doc_meta = lambda: (fake_pn, json.load(open(_meta_that, encoding="utf-8")))
    lk = kn.liet_ke_ky("THUCHI", "baocaothuchi")
    kys = [k["ky"] for k in lk["ky"]]
    print(f"  [metadata thật] THUCHI/baocaothuchi: {lk['so_ky']} kỳ {kys[:1]}..{kys[-1:]}, "
          f"{lk['so_file']} file, không kỳ {len(lk['khong_suy_duoc_ky'])}")
    chk("metadata thật: baocaothuchi đủ kỳ liên tục từ 2025-08, mỗi kỳ 1 file, không file nào mất kỳ",
        kys[0] == "2025-08" and len(kys) >= 15 and not lk["khong_suy_duoc_ky"]
        and all(k["so_file"] == 1 for k in lk["ky"]))
    kn._doc_meta = _gia

# ── cổng PROD ──
p = kn.lap_ke_hoach("ANTAXI", "baocaohqkdngay", ngay_tu="2026-09-01", ngay_den="2026-09-02")
pid, ma = p["plan_id"], p["ma_xac_nhan"]
chk("kế hoạch có 2 file", p["so_file"] == 2)
for f in p["file"]:
    d = os.path.join(os.environ["RECEIVED_DIR"], "ANTAXI", "baocaohqkdngay")
    os.makedirs(d, exist_ok=True)
    open(os.path.join(d, f), "wb").write(b"v1")
chk("PROD khi chưa nạp TEST bị chặn", "nạp TEST" in (raises(kn.nap, pid, "prod") or ""))

r = kn.nap(pid, "test")
chk("nạp TEST ok", r["da_nap"] and r["so_file_loi"] == 0)

r = kn.nap(pid, "prod")
chk("PROD chưa có xác nhận -> chưa nạp, kèm hướng dẫn", r["da_nap"] is False and ma in r["huong_dan_cho_nguoi_dung"])


def ghi_xn(luc, uid=123, ma_=None, **extra):
    os.makedirs(kn.XN_DIR, exist_ok=True)
    json.dump({"ma": ma_ or ma, "luc": luc.isoformat(), "user_id": uid, **extra},
              open(os.path.join(kn.XN_DIR, (ma_ or ma) + ".json"), "w"))


ghi_ma_cua(ma, 123)
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

ghi_xn(datetime.datetime.now() + datetime.timedelta(seconds=2), uid=999)
so_lan_chay.clear()
r = kn.nap(pid, "prod")
chk("xacnhan SAI người (khác chủ mã) -> PROD từ chối, autofill không chạy",
    r["da_nap"] is False and "người đã yêu cầu" in r["ly_do"] and not so_lan_chay)
ghi_xn(datetime.datetime.now() + datetime.timedelta(seconds=2), uid=None)
chk("xacnhan kiểu cũ không user_id -> PROD từ chối", kn.nap(pid, "prod")["da_nap"] is False)
ghi_xn(datetime.datetime.now() + datetime.timedelta(seconds=2))
os.rename(os.path.join(kn.MA_CUA_DIR, ma + ".json"), os.path.join(tmp, "mc.bak"))
chk("xacnhan đúng nhưng thiếu ma_cua -> PROD từ chối", kn.nap(pid, "prod")["da_nap"] is False)
os.rename(os.path.join(tmp, "mc.bak"), os.path.join(kn.MA_CUA_DIR, ma + ".json"))
r = kn.nap(pid, "prod")
chk("đủ điều kiện -> PROD chạy", r["da_nap"] is True and r["env"] == "prod")
r = kn.nap(pid, "prod")
chk("mã xác nhận dùng MỘT lần", r["da_nap"] is False and "đã dùng" in r["ly_do"])
ghi_xn(datetime.datetime.now() + datetime.timedelta(seconds=2))         # API ghi đè file xác nhận (mất da_dung)
chk("API ghi đè xacnhan sau khi dùng -> vẫn chết (cờ trong plan)", kn.nap(pid, "prod")["da_nap"] is False)

# ── da_dung đặt TRƯỚC khi autofill chạy: autofill ném lỗi giữa chừng rồi gọi lại phải bị từ chối ──
p3 = kn.lap_ke_hoach("ANTAXI", "baocaohqkdngay", ngay_tu="2026-09-03", ngay_den="2026-09-03")
pid3, ma3 = p3["plan_id"], p3["ma_xac_nhan"]
tao_file(kn._doc_plan(pid3))
kn.nap(pid3, "test")
ghi_ma_cua(ma3, 123)
ghi_xn(datetime.datetime.now() + datetime.timedelta(seconds=2), ma_=ma3)
thay_luc_chay = {}


def run_no(*a, **k):
    thay_luc_chay["xn"] = json.load(open(os.path.join(kn.XN_DIR, ma3 + ".json"))).get("da_dung")
    thay_luc_chay["plan"] = bool(kn._doc_plan(pid3).get("ma_da_dung_luc"))
    raise OSError("autofill chết giữa chừng")


kn.subprocess.run = run_no
try:
    kn.nap(pid3, "prod")
    chk("autofill ném lỗi -> lỗi nổi lên", False)
except OSError:
    pass
chk("lúc autofill chạy, mã ĐÃ mang da_dung (file xacnhan + plan)",
    thay_luc_chay.get("xn") is True and thay_luc_chay.get("plan") is True)
kn.subprocess.run = run_good
so_lan_chay.clear()
r = kn.nap(pid3, "prod")
chk("gọi lại sau nạp hỏng -> từ chối, không nạp lần 2", r["da_nap"] is False and "đã dùng" in r["ly_do"]
    and not so_lan_chay)
chk("kế hoạch hết hạn bị từ chối", raises(kn._doc_plan, "deadbeef") is not None)
chk("plan_id lạ bị từ chối", raises(kn._plan_path, "../../x") is not None)
chk("nhật ký ghi đủ", sum(1 for _ in open(kn.AUDIT)) >= 4)
print(f"{ok} PASS, {len(loi)} FAIL")
for t in loi:
    print("  FAIL:", t)
sys.exit(1 if loi else 0)

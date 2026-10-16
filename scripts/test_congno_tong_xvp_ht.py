# -*- coding: utf-8 -*-
"""Công nợ TỔNG theo ngày cho Xanh VP (CĐSPS một tầng) và Hưng Thịnh (CĐPS hợp nhất) — 16/10/2026.
Chạy: .venv/bin/python scripts/test_congno_tong_xvp_ht.py
Số trong bài lấy từ file thật ngày 14/10/2026 (XVP: TK 131 dư nợ 39.831.010.614 / dư có 33.469.759;
HT hợp nhất: 131 đầu kỳ 34,473/42,919 · PS 30,08/0 · cuối kỳ 58,427/36,793 tỷ)."""
import datetime
import json
import os
import sys

os.environ.setdefault("BACKEND_PATH", "/home/itadmin/AI_Dashboard_QT/DashBoard_AI/backend")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import derive_hqkd_ngay as d  # noqa: E402

ok = True


def chk(dieu_kien, ten):
    global ok
    if not dieu_kien:
        print("FAIL:", ten)
        ok = False


def pl(f):
    return json.loads(f[6])


# ── XVP: header MỘT tầng, TK con (1311) không được cộng cha lẫn con ──────────────────────────────
xvp = [["Tên đơn vị: Xanh VP"], [None], [datetime.datetime(2026, 10, 14)],
       ["Tk", "Ten Tk", "Du no dau", "Du co dau", "Ps no", "Ps co", "Du no cuoi", "Du co cuoi"],
       ["111", "Tiền mặt", 58677091, 0, 10516967, 0, 69194058, 0],
       ["131", "Phải thu của khách hàng", 39831010614, 33469759, 0, 0, 39831010614, 33469759],
       ["1311", "Con của 131", 39831010614, 0, 0, 0, 39831010614, 0],
       ["331", "Phải trả cho người bán", 59400000, 1735303318, 100, 0, 59400100, 1735303318]]
fs = {f[1]: f for f in d._snap_cdps1tang_congno(xvp)}
chk(set(fs) == {"PTHU_D", "PTRA_D"}, "chỉ 131/331")
chk(fs["PTHU_D"][4] == 39831010614, "PTHU_D = dư Nợ 131 (gộp), không trừ dư Có 33,47tr")
chk(fs["PTRA_D"][4] == 1735303318, "PTRA_D = dư Có 331 (gộp)")
chk(fs["PTHU_D"][2] == "Phải thu khách hàng (tổng)" and fs["PTRA_D"][2] == "Phải trả nhà cung cấp (tổng)",
    "nhãn giống 2 HTX")
chk(abs(pl(fs["PTRA_D"])["du_dau"] - 1.735303318) < 1e-9, "đầu kỳ PTRA")
# ĐK + tăng − giảm = CK trên chiều chính (phía đọc suy ĐK = CK − tăng + giảm)
p = pl(fs["PTHU_D"])
chk(abs(p["du_dau"] + p["ps_tang"] - p["ps_giam"] - 39.831010614) < 1e-9, "đẳng thức ĐK+tăng−giảm=CK")
chk(d._snap_cdps1tang_congno([["x"]]) == [], "sai mẫu -> rỗng, không ném lỗi")

# ── Đối chiếu ngày của sheet cdps1 ───────────────────────────────────────────────────────────────


class _WS:
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, max_row=None, values_only=True):
        return iter(self.rows[:max_row])


class _WB:
    sheetnames = ["CĐSPS"]

    def __init__(self, rows):
        self.ws = _WS(rows)

    def __getitem__(self, k):
        return self.ws


ts = {"cdps1": "CĐSPS"}
chk(d._cdps1_lech_ngay(_WB(xvp), ts, "2026-10-14") is None, "ngày khớp")
chk("ghi 2026-10-14" in (d._cdps1_lech_ngay(_WB(xvp), ts, "2026-10-15") or ""), "ngày lệch bị báo")
chk(d._cdps1_lech_ngay(_WB([["x"]]), ts, "2026-10-14") is not None, "không có ô ngày -> báo, không đoán")
chk(d._cdps1_lech_ngay(_WB(xvp), {}, "2026-10-14") is None, "đơn vị không khai cdps1 -> bỏ qua")

# ── HT: CĐPS hợp nhất, cột dò theo nhãn, loại trừ không dùng ─────────────────────────────────────
ht = [["BẢNG CÂN ĐỐI SỐ PHÁT SINH HỢP NHẤT THÁNG 10/2026"],
      ["Từ ngày 14/10/2026 Đến ngày 14/10/2026 "], [], [],
      ["Số hiệu tài khoản", "Tên tài khoản", "Số dư đầu kỳ hợp nhất", None, "loài trừ đầu kỳ", None,
       "Số phát sinh trong kỳ hợp nhất", None, "loại trừ cuối kỳ", None, "Số dư cuối kỳ sau hợp nhất"],
      [None, None, "Nợ (10)", "Có (11)", None, None, "Nợ (12)", "Có (13)", "nợ", "có", "nợ", "có"],
      ["131", "Phải thu của khách hàng", 34473469990.8, 42919012633.8, None, None, 30079999992, 0, None, None,
       58427469982.8, 36793012633.8],
      ["1311", "Con", 1, 1, None, None, 1, 1, None, None, 1, 1],
      ["331", "Phải trả cho người bán", 4008753832, 329083401401, None, None, 0, 0, None, None,
       4008753832, 329083401401]]
fs, ly = d._snap_cdps_hop_nhat_congno(ht, "2026-10-14")
chk(ly is None and {f[1] for f in fs} == {"PTHU_D", "PTRA_D"}, "HT đọc được 131/331")
fs = {f[1]: f for f in fs}
chk(abs(fs["PTHU_D"][4] - 58427469982.8) < 1e-3, "PTHU_D HT = dư Nợ cuối kỳ SAU hợp nhất")
chk(abs(fs["PTRA_D"][4] - 329083401401) < 1e-3, "PTRA_D HT = dư Có 331")
p = pl(fs["PTHU_D"])
chk(abs(p["du_dau"] - 34.47346999) < 1e-6 and abs(p["ps_tang"] - 30.079999992) < 1e-9, "đầu kỳ/tăng HT")
chk(abs(p["du_dau"] + p["ps_tang"] - p["ps_giam"] - 58.4274699828) < 1e-6, "đẳng thức HT")
# kỳ sai / luỹ kế
chk(d._snap_cdps_hop_nhat_congno(ht, "2026-10-15")[0] == [], "kỳ không phải ngày ảnh chụp -> bỏ")
luy = [list(r) for r in ht]
luy[1] = ["Từ ngày 01/10/2026 Đến ngày 14/10/2026 "]
fl, _ = d._snap_cdps_hop_nhat_congno(luy, "2026-10-14")
chk(fl and abs(fl[0][4] - 58427469982.8) < 1e-3 and pl(fl[0])["ps_tang"] == 0.0 and pl(fl[0])["du_dau"] == 0.0,
    "bản luỹ kế giữ số dư, bỏ phát sinh/đầu kỳ")
bad = [list(r) for r in ht]
bad[4] = ["TK", "Tên"]
chk(d._snap_cdps_hop_nhat_congno(bad, "2026-10-14")[1] is not None, "thiếu header -> nói lý do")

# ── Cấu hình ─────────────────────────────────────────────────────────────────────────────────────
chk(d._UNITS["XANHVINHPHUC"]["sheet_sodu"].get("cdps1") == "CĐSPS", "XVP khai cdps1")
chk("cdps" not in d._UNITS["XANHVINHPHUC"]["sheet_sodu"], "XVP KHÔNG khai cdps hai tầng")
chk(d._UNITS["HUNGTHINH"]["sheet_sodu"].get("cdps_hn"), "HT khai cdps_hn")
print("PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)

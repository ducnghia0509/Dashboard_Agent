# -*- coding: utf-8 -*-
"""NỐI CỤM SỐ DƯ THEO NGÀY cho Showroom Vinfast (SRVF) và Xưởng dịch vụ (XDV) — nguồn TỰ ĐỘNG Cyber.

VÌ SAO KHÔNG ĐỌC EXCEL LẠI: bộ file ngày của hai khối này ĐÃ được `cron_srvf_daily` /
`cron_xdv_daily` bóc vào `raw_rows` hằng ngày qua các spec `vhkd_cdkt_ngay` / `vhkd_cdps_ngay` /
`vhkd_pthu_hoadon_ngay` / `xdv_congno_ro_ngay`, và ĐỦ MỌI CỘT cần dùng:

    VHKD_CDKT_D   dim1 = mã số CĐKT · amount = Số cuối kỳ · amount2 = Số đầu kỳ
                  payload: chi_tieu · so_dau_nam · tai_khoan · cong_thuc
    VHKD_CDPS_D   dim1 = số hiệu TK · dim2 = cấp TK · amount = PS Nợ · amount2 = PS Có
                  payload: ten_tk · no_dau_ky · co_dau_ky · du_no_cuoi · du_co_cuoi
    VHKD_PTHU_HD  công nợ phải thu CHI TIẾT theo khách/hợp đồng (dim1 = trong hạn / quá hạn)
    XDV_CN_RO_D   công nợ phải thu XDV chi tiết theo R/O (payload: ma_doi_tuong · ten_doi_tuong)

Cả bốn spec đều tự gắn cờ `"chua_len_man": "1"` — chính chúng ghi rằng "nạp để sẵn, chưa khai ô
đích". File này LÀ ô đích: đổi sang đúng report_type + đúng quy ước dim/payload mà màn TCKT đọc,
không đụng gì tới khâu bóc file. Đọc lại Excel ở đây là chép đôi một phép bóc đã được kiểm chứng
và đang chạy hằng ngày.

RA ĐƯỢC NHỮNG GÌ
    SRVF   BS_D · TS_D · TSNV_D  (từ CĐKT)      -> màn Tài sản-Nguồn vốn, Tài sản
           THUE_D · HH_D         (từ CĐPS)      -> màn Thuế, Tồn kho
           PTHU_D · PTRA_D       (từ CĐPS + chi tiết khách của VHKD_PTHU_HD) -> màn Công nợ
    XDV    PTHU_D                (từ XDV_CN_RO_D, chi tiết theo đối tượng)   -> màn Công nợ

    XDV KHÔNG có CĐKT/CĐPS toàn phần ở nguồn -> không dựng được TS-Nguồn vốn · Tài sản · Thuế ·
    Tồn kho. Đó là thiếu NGUỒN, không phải thiếu code: phải xin kế toán xuất thêm như SR.

⚠ BẢNG CÂN ĐỐI CỦA SR CHƯA CÂN. Đo 18/09/2026 trên 6 ngày gần nhất: mã 270 (tổng tài sản) và 440
(tổng nguồn vốn) lệch từ 4,6 đến 500,6 tỷ, và 15->16/09 tổng tài sản nhảy 2.931 -> 14.330 tỷ do
kế toán sửa template giữa chừng. `extract_specs/vhkd_cdkt_ngay.json` mục `_bay` đã cảnh báo sẵn
bằng chữ in hoa. User chốt ngày 18/09/2026 "lên hết tất cả" nên vẫn nối; màn Tài sản-Nguồn vốn tự
có ô "LỆCH x tỷ" cạnh thẻ tổng nên số lệch HIỆN RA chứ không bị giấu. Đừng lấy màn này ra số công
bố khi ô đó chưa về 0.

QUY ƯỚC dim/payload GIỮ ĐÚNG BẢN THÁNG + bản Trạm sạc (`derive_hqkd_ngay._snap_*_facts`) — lệch
quy ước là cùng một khoản mục hiện hai bộ tên ở hai chế độ. Xem từng hàm `_facts_*` bên dưới.

CHỌN ĐÚNG MỘT TẦNG TÀI KHOẢN: `VHKD_CDPS_D` chứa cả Cấp 1..4 (131 và 1311, 1316, 1318...). Khớp
theo TIỀN TỐ như bản đọc-file sẽ cộng cha lẫn con = gấp đôi. Ở đây khớp `dim1` BẰNG ĐÚNG mã TK
đích, mỗi mã trong file chỉ có một dòng.

Chạy:  .venv/bin/python scripts/derive_sodu_tudong_ngay.py --period 2026-09 [--dry-run]
"""
import argparse
import datetime
import json
import os
import sys

import psycopg
from psycopg.rows import dict_row, tuple_row

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [p for p in (os.path.dirname(_HERE), _HERE) if p not in sys.path]
from servers.common import dataset_ky as _DSK  # noqa: E402

# `--env` thay vì đặt DATABASE_URL ngay trong dòng crontab (sửa 19/09/2026). Dòng cron cũ ghi
# `DATABASE_URL="postgresql://tc:tc_%24production@..."` và **không bao giờ chạy**: trong crontab,
# `%` là ký tự ĐẶC BIỆT (biến thành xuống dòng, phần sau thành stdin), nên lệnh bị cắt ngang ở
# `tc_` — không có lỗi nào hiện ra, chỉ là file log không bao giờ được tạo. Các cron khác trong
# nhà đều dùng `--env`, đi theo cho đồng nhất và hết hẳn chuyện escape.
MOI_TRUONG = {
    "test": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
    "prod": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard",
}
DB_URL = (os.environ.get("DATABASE_URL") or os.environ.get("TC_DATABASE_URL")
          or "postgresql://tc:tc@localhost:5433/tc_dashboard")

SR_KHOI, XDV_KHOI, CONG_TY = "Khối KD Vinfast - Showroom", "Khối KD Vinfast - XDV", "TC"

# Khoá idempotent THEO KỲ, không theo file: mỗi lượt dựng lại toàn bộ ngày của kỳ từ nhiều
# report_type nguồn, cùng lý do như `derive_hqkd_ngay.period_source_key`.
def _key(nguon, period):
    return f"{nguon}::sodu_tudong::{period}"


# ── Bảng tra, lấy NGUYÊN VĂN từ `derive_hqkd_ngay` để hai đường ra cùng một bộ nhãn ────────────
_CDKT_BS = ("270", "300", "400")
_CDKT_TS = {"gtcl": "221", "nguyen_gia": "222", "hao_mon": "223"}
_TSNV_NHOM = {**{m: "TS ngắn hạn" for m in ("110", "120", "130", "140", "150")},
              **{m: "TS dài hạn" for m in ("210", "220", "230", "240", "250", "260")},
              **{m: "Nợ phải trả" for m in ("310", "330")}, "400": "Vốn chủ"}
_THUE_TK = {"133": ("Phải thu", "Thuế GTGT được khấu trừ", "no"),
            "3331": ("Phải nộp", "Thuế Giá trị gia tăng (GTGT)", "co"),
            "3334": ("Phải nộp", "Thuế Thu nhập doanh nghiệp (TNDN)", "co"),
            "3335": ("Phải nộp", "Thuế Thu nhập cá nhân (TNCN)", "co"),
            "3338": ("Phải nộp", "Thuế, phí khác", "co")}
_HH_TK = {"151": "Hàng mua đang đi đường", "152": "Nguyên liệu, vật liệu",
          "153": "Công cụ, dụng cụ", "154": "Chi phí sản xuất kinh doanh dở dang",
          "155": "Thành phẩm", "156": "Hàng hóa"}


def _pl(**kv):
    return json.dumps({**kv, "unit": "ty", "grain": "day"}, ensure_ascii=False)


def _so(x):
    return x if isinstance(x, (int, float)) else 0.0


def _doc(cur, rt, period, ds_id):
    """{ngày: [dòng]} của một report_type nguồn trong kỳ. Dòng đã giải payload sẵn.

    Bó theo `dataset_id` chứ không chỉ `period_month`: hai dataset cùng kỳ (kỳ bị khai sinh lại,
    hoặc bản nhập tay song song) sẽ trộn dòng của nhau và số nhân đôi mà không có dấu hiệu gì."""
    cur.execute(
        "SELECT ngay, dim1, dim2, dim3, amount, amount2, payload FROM raw_rows "
        "WHERE report_type=%s AND period_month=%s AND dataset_id=%s AND ngay IS NOT NULL",
        (rt, period, ds_id))
    theo_ngay = {}
    for r in cur.fetchall():
        try:
            r["pl"] = json.loads(r["payload"] or "{}")
        except Exception:
            r["pl"] = {}
        theo_ngay.setdefault(r["ngay"], []).append(r)
    return theo_ngay


# ── CĐKT -> BS_D / TS_D / TSNV_D ────────────────────────────────────────────────────────────────
def _facts_cdkt(rows):
    """dòng VHKD_CDKT_D của MỘT ngày -> [(rt, dim1, dim2, dim3, amount_tỷ, payload)].

    `amount` của nguồn đã là Số cuối kỳ (tỷ) — đúng thứ cụm số dư cần. `amount2` (Số đầu kỳ) KHÔNG
    ghi vào cột `amount2` của dòng đích: phía đọc suy đầu kỳ từ `_rows_dau_ky_ngay` (số dư ngày sớm
    nhất trong khoảng người xem chọn), lấy đầu kỳ của riêng ngày sẽ lệch với ba màn còn lại.
    """
    byma = {}
    for r in rows:
        ma = (r["dim1"] or "").strip()
        if ma.isdigit() and len(ma) == 3 and ma not in byma:
            byma[ma] = (_so(r["amount"]), (r["pl"].get("chi_tieu") or "").strip())
    out = []
    for ma in _CDKT_BS:
        if ma in byma:
            out.append(("BS_D", ma, None, ma, byma[ma][0], None))
    if _CDKT_TS["gtcl"] in byma:
        ng = byma.get(_CDKT_TS["nguyen_gia"], (0.0, ""))[0]
        hm = abs(byma.get(_CDKT_TS["hao_mon"], (0.0, ""))[0])
        out.append(("TS_D", "TSCĐ (theo CĐKT)", None, "TSCĐ (theo CĐKT)", byma[_CDKT_TS["gtcl"]][0],
                    _pl(nguyen_gia=ng, hao_mon=hm, khau_hao_ky=0.0,
                        het_kh_con_sd=0.0, het_kh_thanh_ly=0.0)))
    # TSNV = TOÀN BỘ dòng CĐKT cho bảng cơ cấu. `dim1` CHỈ gán cho dòng cấp La Mã + mã 400, y hệt
    # bản tháng — gán rộng hơn là dòng cha và dòng con cùng lọt một nhóm rồi cộng đôi.
    for ma, (v, ten) in byma.items():
        if not ten:
            continue
        out.append(("TSNV_D", _TSNV_NHOM.get(ma, ""), ten, "", v,
                    _pl(ps_tang=0.0, ps_giam=0.0, ma_so=ma)))
    return out


# ── CĐPS -> THUE_D / HH_D / PTHU_D / PTRA_D (tổng theo tài khoản) ───────────────────────────────
def _facts_cdps(rows):
    """dòng VHKD_CDPS_D của MỘT ngày -> facts. Khớp `dim1` BẰNG ĐÚNG mã TK (xem docstring module)."""
    bytk = {}
    for r in rows:
        tk = (r["dim1"] or "").strip()
        if tk and tk not in bytk:
            bytk[tk] = r
    out = []

    def bon_so(r):
        """(dư nợ đầu, dư có đầu, PS nợ, PS có, dư nợ cuối, dư có cuối) — đơn vị tỷ."""
        return (_so(r["pl"].get("no_dau_ky")), _so(r["pl"].get("co_dau_ky")),
                _so(r["amount"]), _so(r["amount2"]),
                _so(r["pl"].get("du_no_cuoi")), _so(r["pl"].get("du_co_cuoi")))

    for tk, (nhom, ten, chieu) in _THUE_TK.items():
        r = bytk.get(tk)
        if not r:
            continue
        dn, dc, pn, pc, cn, cc = bon_so(r)
        dau, cuoi = (dn - dc, cn - cc) if chieu == "no" else (dc - dn, cc - cn)
        tang, giam = (pn, pc) if chieu == "no" else (pc, pn)
        if not (dau or tang or giam or cuoi):
            continue
        out.append(("THUE_D", nhom, ten, nhom, cuoi,
                    _pl(du_dau=dau, ps_tang=tang, ps_giam=giam)))
    for tk, ten in _HH_TK.items():
        r = bytk.get(tk)
        if not r:
            continue
        dn, dc, pn, pc, cn, cc = bon_so(r)
        dau, cuoi = dn - dc, cn - cc
        if not (dau or pn or pc or cuoi):
            continue
        out.append(("HH_D", ten, None, ten, cuoi,
                    _pl(du_dau=dau, tk=tk, nhap=pn, xuat=pc, cham_lc=0.0, ton_3m=0.0, ton_36=0.0)))
    return out, bytk


# ── Công nợ CHI TIẾT theo đối tượng ─────────────────────────────────────────────────────────────
def _facts_congno(rows, rt, chieu, ten_key, ma_key):
    """Gom dòng chi tiết theo ĐỐI TƯỢNG -> facts PTHU_D/PTRA_D.

    Quy ước giống `_snap_congno_facts`: dim1 = dim3 = tên đối tượng · amount = số dư RÒNG ·
    payload mang du_dau/ps_tang/ps_giam/ma_dt. Bảng Top 10 của màn Công nợ đọc đúng bộ này.
    """
    gom = {}
    for r in rows:
        ten = (r["pl"].get(ten_key) or "").strip()
        if not ten:
            continue
        a = gom.setdefault(ten, {"dau": 0.0, "tang": 0.0, "giam": 0.0, "cuoi": 0.0,
                                 "ma": (r["pl"].get(ma_key) or "").strip()})
        dn, dc = _so(r["pl"].get("no_dau_ky")), _so(r["pl"].get("co_dau_ky"))
        pn, pc = _so(r["pl"].get("ps_no")), _so(r["pl"].get("ps_co"))
        cuoi = _so(r["amount"])
        if chieu == "no":
            a["dau"] += dn - dc; a["tang"] += pn; a["giam"] += pc; a["cuoi"] += cuoi
        else:
            a["dau"] += dc - dn; a["tang"] += pc; a["giam"] += pn; a["cuoi"] += cuoi
    return [(rt, ten, None, ten, v["cuoi"],
             _pl(du_dau=v["dau"], ps_tang=v["tang"], ps_giam=v["giam"], ma_dt=v["ma"]))
            for ten, v in gom.items()
            if any(abs(v[k]) > 1e-9 for k in ("dau", "tang", "giam", "cuoi"))]


def _facts_congno_tong(bytk, tk, rt, chieu, ten):
    """Không có nguồn chi tiết thì ít nhất ra MỘT dòng tổng theo tài khoản — thẻ tổng của màn Công
    nợ vẫn đúng, chỉ thiếu bảng Top 10. Trung thực hơn là để trống cả màn."""
    r = bytk.get(tk)
    if not r:
        return []
    dn, dc = _so(r["pl"].get("no_dau_ky")), _so(r["pl"].get("co_dau_ky"))
    pn, pc = _so(r["amount"]), _so(r["amount2"])
    cn, cc = _so(r["pl"].get("du_no_cuoi")), _so(r["pl"].get("du_co_cuoi"))
    dau, cuoi = (dn - dc, cn - cc) if chieu == "no" else (dc - dn, cc - cn)
    tang, giam = (pn, pc) if chieu == "no" else (pc, pn)
    if not (dau or tang or giam or cuoi):
        return []
    return [(rt, ten, None, ten, cuoi, _pl(du_dau=dau, ps_tang=tang, ps_giam=giam, ma_dt=tk))]


RT_RA = ("BS_D", "TS_D", "TSNV_D", "THUE_D", "HH_D", "PTHU_D", "PTRA_D")


def dung(period, cur, ds_id):
    """-> {nguồn: {ngày: [facts]}}. Không đụng DB ghi, để `main` quyết dry-run."""
    cdkt = _doc(cur, "VHKD_CDKT_D", period, ds_id)
    cdps = _doc(cur, "VHKD_CDPS_D", period, ds_id)
    pthu_hd = _doc(cur, "VHKD_PTHU_HD", period, ds_id)
    xdv_ro = _doc(cur, "XDV_CN_RO_D", period, ds_id)

    sr = {}
    for ngay in sorted(set(cdkt) | set(cdps)):
        fs = list(_facts_cdkt(cdkt.get(ngay, [])))
        bytk = {}
        if ngay in cdps:
            f2, bytk = _facts_cdps(cdps[ngay])
            fs += f2
        # Phải THU: ưu tiên chi tiết theo khách (VHKD_PTHU_HD); không có thì lấy tổng TK 131.
        ct = _facts_congno(pthu_hd.get(ngay, []), "PTHU_D", "no", "ten_khach", "ma_khach")
        fs += ct or _facts_congno_tong(bytk, "131", "PTHU_D", "no", "Phải thu khách hàng (tổng)")
        # Phải TRẢ: SR không có nguồn chi tiết nhà cung cấp -> chỉ tổng TK 331.
        fs += _facts_congno_tong(bytk, "331", "PTRA_D", "co", "Phải trả nhà cung cấp (tổng)")
        if fs:
            sr[ngay] = fs

    xdv = {}
    for ngay, rows in xdv_ro.items():
        fs = _facts_congno(rows, "PTHU_D", "no", "ten_doi_tuong", "ma_doi_tuong")
        if fs:
            xdv[ngay] = fs
    return {"TEST_SR": (SR_KHOI, sr), "TEST_XDV": (XDV_KHOI, xdv)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--period", help="YYYY-MM; bỏ trống = tháng hiện tại theo giờ VN")
    ap.add_argument("--env", choices=tuple(MOI_TRUONG), help="chọn DB; không khai thì theo DATABASE_URL")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    # Mặc định lấy tháng theo GIỜ VN, không phải UTC: chạy lúc 01:45 VN (18:45 UTC hôm trước) mà
    # tính theo UTC thì ngày 01 hằng tháng sẽ nạp nhầm kỳ của tháng trước.
    if not a.period:
        a.period = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))).strftime("%Y-%m")
    db_url = MOI_TRUONG[a.env] if a.env else DB_URL

    conn = psycopg.connect(db_url, row_factory=dict_row)
    try:
        cur = conn.cursor()
        cur_t = conn.cursor(row_factory=tuple_row)
        dataset_id, _tt = _DSK.lay_hoac_tao_ky(cur_t, a.period, nguon="derive_sodu_tudong_ngay")
        if not dataset_id:
            print(json.dumps({"ok": False, "error": _DSK.giai_thich(a.period, _tt)},
                             ensure_ascii=False))
            return
        ra = dung(a.period, cur, dataset_id)
        tom = {}
        for nguon, (khoi, theo_ngay) in ra.items():
            dem = {}
            for ngay, fs in theo_ngay.items():
                for f in fs:
                    dem[f[0]] = dem.get(f[0], 0) + 1
            tom[nguon] = {"so_ngay": len(theo_ngay), "ngay_cuoi": max(theo_ngay) if theo_ngay else None,
                          "dong_theo_loai": dem}
        if a.dry_run:
            print(json.dumps(tom, ensure_ascii=False, indent=2))
            return

        ghi = 0
        for nguon, (khoi, theo_ngay) in ra.items():
            src = _key(nguon, a.period)
            cur.execute("DELETE FROM raw_rows WHERE source_file=%s", (src,))
            recs, i = [], 0
            for ngay in sorted(theo_ngay):
                for rt, dim1, dim2, dim3, amount, payload in theo_ngay[ngay]:
                    i += 1
                    recs.append((dataset_id, rt, 8900000 + i, ngay, CONG_TY, khoi, None,
                                 a.period, amount, None, dim1, dim2, dim3,
                                 payload or _pl(), src))
            if recs:
                cur.executemany(
                    "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi,"
                    " cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload,"
                    " source_file) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", recs)
            ghi += len(recs)
        conn.commit()
        print(json.dumps({"ok": True, "period": a.period, "written": ghi, **tom},
                         ensure_ascii=False, indent=2))
    finally:
        conn.close()


if __name__ == "__main__":
    main()

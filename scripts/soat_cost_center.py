# -*- coding: utf-8 -*-
"""SOÁT COST CENTER của báo cáo NGÀY — dò cột mới + CHỨNG MINH chỗ lệch do đâu.

VÌ SAO CÓ FILE NÀY (18/09/2026). Layout "duan" khớp cột dự án bằng danh sách từ khoá gõ cứng
`derive_hqkd_ngay._CC_DUAN`. Kế toán thêm một dự án mới = thêm một cột Excel, mà danh sách đó
không có từ khoá tương ứng -> cột **rơi im lặng**: tiền có trong file, không lên dashboard, không
một dòng cảnh báo nào. "Bình Phước" từng nằm chờ như vậy.

Module này KHÔNG ghi gì vào DB và KHÔNG đổi cách tính số nào — chỉ đọc, đếm, rồi kết luận. Chạy
được độc lập để người soát, và `cron_hqkdngay_daily` gọi để xếp việc cho admin duyệt.

BA CÂU HỎI NÓ TRẢ LỜI, tách bạch — vì gộp lại là đúng cái bẫy đã mất một lượt phân tích ngày
18/09: thấy lệch 341tr rồi vội kết luận "thiếu cost center", trong khi cột Tân Thịnh VẪN CÓ và
vẫn được đọc, chỉ là nguồn bỏ trống 30/31 ngày. Hai nguyên nhân đó đòi hai người khác nhau đi
sửa (admin duyệt mã mới / kế toán nhập bù số), nên tuyệt đối không được trộn:

  1. CỘT CHƯA CÓ MÃ  -> cột có tên, có tiền, không khớp từ khoá nào. Việc của ADMIN: duyệt mã.
  2. CỘT CÓ MÃ NHƯNG NGUỒN THIẾU SỐ -> đọc đủ rồi mà Σ ngày vẫn hụt so với BCTC tháng.
     Việc của KẾ TOÁN đơn vị: nhập bù.
  3. CỘT BỊ BỎ CÓ ĐIỀU KIỆN -> "Tổng Dự án"/"HO Dự án" cố ý không đọc. Chỉ im lặng khi nó = 0
     hoặc TRÙNG KHÍT một cột khác (ca công thức hỏng); mang số riêng thật thì phải kêu.

CHUẨN ĐỐI CHIẾU LÀ BCTC THÁNG, KHÔNG PHẢI CỘT "TỔNG DỰ ÁN" CỦA FILE. Đo 18/09 trên chính file
T8: cột Tổng lệch Σ7 cột dự án tới 879.126.946 đ — ngày 28/29/30 cộng đôi Cao Bằng qua cột HO
hỏng công thức, ngày 31 vênh 621.080.352 đ không cột nào giải thích. Lấy cột Tổng làm chuẩn là
tự buộc mình vào một con số sai.
"""
import json
import os
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl                                          # noqa: E402
import derive_hqkd_ngay as D                             # noqa: E402

# Cột cố ý KHÔNG đọc: là cột TỔNG HỢP, đọc vào là cộng đôi với chính các cột con của nó.
_COT_BO = ("tong du an", "ho du an")
_NGUONG = 1.0            # đồng — dưới ngưỡng này coi là sai số làm tròn, không phải lệch


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _ma_cua(ten: str):
    """Tên cột -> mã cost center. None = CHƯA CÓ MÃ.

    GỌI THẲNG `D._ma_cost_center` chứ KHÔNG tự tra lại `D._CC_DUAN`: bản đồ có HAI nguồn (danh
    sách trong code + bảng `cost_center_map` admin đã duyệt), tra thiếu nguồn thứ hai là cột vừa
    được duyệt vẫn bị báo "chờ duyệt" ở mọi lượt sau -> chuông không bao giờ tắt và admin bấm
    duyệt mãi không hết việc. Đã dính đúng lỗi này khi nghiệm thu 18/09.
    """
    return D._ma_cost_center(ten)


def doc_cot(path: str) -> dict:
    """Quét mọi sheet ngày -> {ten_cot: {ma, tong, so_ngay}} + cột tổng của file, để đối chiếu."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    cot, tong_file = {}, 0.0
    for sn in wb.sheetnames:
        if not sn.strip().isdigit():
            continue
        rows = list(wb[sn].iter_rows(min_row=1, max_row=90, max_col=20, values_only=True))
        hdr = next((r for r in rows[:10]
                    if any(D._nd(c) == "tong du an" for c in r if c is not None)), None)
        dtt = next((r for r in rows
                    if D._nd(r[1] if len(r) > 1 else "") == "doanh thu thuan"), None)
        if hdr is None or dtt is None:
            continue
        for j, c in enumerate(hdr):
            if not isinstance(c, str) or not c.strip() or j >= len(dtt):
                continue
            ten, v = c.strip(), _f(dtt[j])
            if D._nd(ten) == "tong du an":
                tong_file += v
                continue
            e = cot.setdefault(ten, {"ma": _ma_cua(ten), "tong": 0.0, "so_ngay": 0,
                                     "theo_ngay": {},
                                     "bo_co_chu_y": D._nd(ten) in _COT_BO})
            e["tong"] += v
            e["so_ngay"] += 1 if abs(v) > _NGUONG else 0
            e["theo_ngay"][sn] = v
    wb.close()
    return {"cot": cot, "tong_cot_tong_cua_file": tong_file}


def _trung_cot_khac(ten: str, cot: dict) -> str:
    """Cột bỏ-có-điều-kiện có phải BẢN SAO của một cột khác không (ca công thức hỏng)?

    So THEO TỪNG NGÀY, không so tổng tháng: cột "HO Dự án" của T8 chỉ sao chép Cao Bằng ở ngày
    28/29/30, các ngày khác bằng 0 — tổng tháng vì thế không trùng tổng cột nào, so bằng tổng sẽ
    kết luận nhầm là "số riêng" rồi báo động oan. Bản sao = MỌI ngày cột này khác 0 đều khớp khít
    cùng một cột khác trong CHÍNH ngày đó.
    """
    ngay_co = {d: v for d, v in cot[ten]["theo_ngay"].items() if abs(v) > _NGUONG}
    if not ngay_co:
        return None
    for k, v in cot.items():
        if k == ten or v["bo_co_chu_y"]:
            continue
        if all(abs(v["theo_ngay"].get(d, 0.0) - x) <= _NGUONG for d, x in ngay_co.items()):
            return k
    return None


def doi_chieu(path: str, thang_theo_cc: dict) -> dict:
    """So Σ ngày với BCTC THÁNG theo TỪNG cost center rồi quy trách nhiệm cho đúng người.

    `thang_theo_cc`: {ma_cost_center: doanh_thu_thuan_dong} lấy từ report_type DTHU/HQKD của
    BCTC tháng — NGUỒN ĐỘC LẬP với file ngày, đó là lý do nó đủ tư cách làm chuẩn.
    """
    q = doc_cot(path)
    cot = q["cot"]
    ngay_theo_cc = {}
    for ten, e in cot.items():
        if e["bo_co_chu_y"] or not e["ma"]:
            continue
        ngay_theo_cc[e["ma"]] = ngay_theo_cc.get(e["ma"], 0.0) + e["tong"]

    chua_co_ma = [{"ten_cot": t, "tong_dong": e["tong"], "so_ngay_co_so": e["so_ngay"]}
                  for t, e in cot.items()
                  if not e["bo_co_chu_y"] and not e["ma"] and abs(e["tong"]) > _NGUONG]

    # Cột bỏ-có-điều-kiện: chỉ được im khi = 0 hoặc trùng khít cột khác.
    bo_bat_thuong = []
    for t, e in cot.items():
        if not e["bo_co_chu_y"] or abs(e["tong"]) <= _NGUONG:
            continue
        ban_sao = _trung_cot_khac(t, cot)
        if ban_sao:
            bo_bat_thuong.append({"ten_cot": t, "tong_dong": e["tong"], "loai": "ban_sao",
                                  "trung_voi": ban_sao})
        else:
            bo_bat_thuong.append({"ten_cot": t, "tong_dong": e["tong"], "loai": "so_rieng"})

    # KHÔNG đối chiếu khi CHƯA CÓ BCTC THÁNG của kỳ đó. BCTC tháng không có cron, thường về sau
    # ngày 10 (xem memory bctc-thang-khong-co-cron-keo-tay), nên đầu mỗi tháng `thang_theo_cc` rỗng
    # -> mọi cost center bị so với 0 và báo "lệch âm" toàn bộ. Đó là BÁO ĐỘNG GIẢ thuần tuý: không
    # có gì sai cả, chỉ là vế để so chưa tồn tại. Bắt được lúc nghiệm thu 18/09 với kỳ 2026-09
    # (CB_DA -3,26 tỷ / PQ_DA -3,77 tỷ / TT_DA -88 tr — không cái nào là lệch thật).
    ket = []
    if not thang_theo_cc:
        return {"cot_chua_co_ma": chua_co_ma, "cot_bo_bat_thuong": bo_bat_thuong,
                "doi_chieu_theo_cost_center": [],
                "bo_qua_doi_chieu": "chưa có BCTC tháng của kỳ này để đối chiếu",
                "tong_cot_tong_cua_file": q["tong_cot_tong_cua_file"],
                "tong_sum_ngay_da_doc": sum(ngay_theo_cc.values()),
                "can_admin_duyet": bool(chua_co_ma)}
    for cc in sorted(set(ngay_theo_cc) | set(thang_theo_cc)):
        ng, th = ngay_theo_cc.get(cc, 0.0), float(thang_theo_cc.get(cc, 0.0))
        lech = th - ng
        if abs(lech) <= _NGUONG:
            continue
        so_ngay = sum(e["so_ngay"] for e in cot.values() if e["ma"] == cc)
        co_cot = any(e["ma"] == cc for e in cot.values())
        khop = next((c for c in chua_co_ma if abs(c["tong_dong"] - lech) <= _NGUONG), None)
        if not co_cot and khop:
            loai, ai, vi = "THIEU_COST_CENTER", "admin", \
                f"cột '{khop['ten_cot']}' chưa có mã, tổng khớp KHÍT chỗ lệch"
        elif not co_cot:
            loai, ai, vi = "KHONG_CO_COT", "ketoan", "BCTC tháng có số mà file ngày không có cột"
        else:
            loai, ai, vi = "NGUON_THIEU_SO_NGAY", "ketoan", \
                f"cột CÓ và đã đọc, nhưng chỉ {so_ngay} ngày có số"
        ket.append({"cost_center": cc, "sum_ngay": ng, "bctc_thang": th, "lech_dong": lech,
                    "loai": loai, "ai_xu_ly": ai, "vi_sao": vi, "so_ngay_co_so": so_ngay})

    return {"cot_chua_co_ma": chua_co_ma, "cot_bo_bat_thuong": bo_bat_thuong,
            "doi_chieu_theo_cost_center": ket,
            "tong_cot_tong_cua_file": q["tong_cot_tong_cua_file"],
            "tong_sum_ngay_da_doc": sum(ngay_theo_cc.values()),
            "can_admin_duyet": bool(chua_co_ma)}


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--thang-json", help='{"CB_DA": 123, ...} doanh thu thuần BCTC tháng')
    a = ap.parse_args()
    th = json.loads(a.thang_json) if a.thang_json else {}
    print(json.dumps(doi_chieu(a.path, th), ensure_ascii=False, indent=2))


# ────────────────────────── XẾP VIỆC CHO ADMIN DUYỆT ──────────────────────────
# Tái dùng chuông 🔔 `source_change_events` (migration 0033) thay vì dựng hàng đợi mới: admin đã
# quen bấm duyệt ở đó, và nó có sẵn chống trùng theo `dedupe_key` + nút "Đã xử lý".
# KHÔNG dùng bảng `approvals` — `approvals_type_check` chỉ cho payment/contract/purchase/hr.

_LAYOUT_MAC_DINH = "duan"

# Danh mục mà một cost center MỚI cần có để lên được Ô LỌC, không chỉ lên số: nhãn đọc được +
# khối + pháp nhân. Ô chọn "Cost center" và nhãn tiếng Việt đọc từ danh mục TĨNH
# (`master_data.json::costCenters`), không đọc từ dữ liệu — thiếu ba thứ này thì dự án mới có
# tiền trong DB mà vẫn không lọc được và hiện ra mã trần (đo thật 18/09/2026 với BINHPHUOC_DA).
# `khoi` PHẢI khớp đúng chuỗi trong `master_data.json::khoi`, lệch một chữ là bộ lọc theo Khối bỏ sót.
# `khoi`/`cong_ty` lấy Y HỆT `derive_hqkd_ngay._UNITS` — lệch một chữ là bộ lọc theo Khối bỏ sót.
# `bo_tien_to`: tên cột đã mang sẵn tiền tố đơn vị ("Showroom Hạ Long", "XDV Việt Trì") nên cắt đi
# trước khi sinh mã, nếu không ra `SHOWROOMHALONG_SR` dài vô ích.
_DANH_MUC_LAYOUT = {
    "duan":   {"tien_to_ten": "Dự án ", "khoi": "Khối KD Dự án", "cong_ty": "TC",
               "hau_to": "_DA", "bo_tien_to": ("du an",)},
    "srvf":   {"tien_to_ten": "", "khoi": "Khối KD Vinfast - Showroom", "cong_ty": "TC",
               "hau_to": "_SR", "bo_tien_to": ("showroom", "vinfast")},
    "xdv":    {"tien_to_ten": "", "khoi": "Khối KD Vinfast - XDV", "cong_ty": "TC",
               "hau_to": "_XDV", "bo_tien_to": ("xdv", "vf", "chi nhanh vinfast")},
    "antaxi": {"tien_to_ten": "Depot ", "khoi": "Khối KD Dịch vụ An Taxi", "cong_ty": "AAG",
               "hau_to": "_AT", "bo_tien_to": ()},
}


def ma_de_xuat(ten_cot: str, layout: str = _LAYOUT_MAC_DINH) -> str:
    """'Bình Phước' + layout duan -> 'BINHPHUOC_DA'. Chỉ là ĐỀ XUẤT để admin sửa/duyệt.

    Hậu tố PHẢI theo layout: `_DA`/`_SR`/`_XDV`/`_AT` đúng quy ước mã đang dùng trong
    `master_data.json`. Gắn cứng `_DA` cho mọi layout là đề xuất mã sai đơn vị ngay từ đầu.
    """
    dm = _DANH_MUC_LAYOUT.get(layout, {})
    n = D._nd(ten_cot)
    for tt in dm.get("bo_tien_to", ()):
        if n.startswith(tt):
            n = n[len(tt):].strip()
    n = "".join(ch for ch in n.upper() if ch.isalnum())
    return f"{n[:24]}{dm.get('hau_to', '_DA')}" if n else ""


def thang_theo_cost_center(conn, source_file_thang: str) -> dict:
    """{ma_cc: doanh_thu_thuan_dong} từ BCTC THÁNG — nguồn ĐỘC LẬP với file ngày."""
    rows = conn.execute(
        "SELECT cost_center, SUM(amount)*1e9 FROM raw_rows "
        "WHERE source_file=%s AND report_type='DTHU' AND dim1='Doanh thu thuần' "
        "GROUP BY cost_center", (source_file_thang,)).fetchall()
    return {cc: float(v or 0) for cc, v in rows if cc}


def xep_viec(conn, path: str, period: str, company: str, source_file_thang: str = None,
             layout: str = _LAYOUT_MAC_DINH, log=print) -> dict:
    """Soát 1 file ngày rồi ghi việc CHỜ DUYỆT. Trả chính báo cáo soát.

    KHÔNG ném lỗi ra ngoài: đây là lớp giám sát chạy kèm lượt nạp, hỏng nó mà làm hỏng luôn
    lượt nạp thì lợi bất cập hại — số liệu vẫn phải vào DB.
    """
    try:
        if layout == _LAYOUT_MAC_DINH:
            # "duan" giữ đường CŨ: ngoài dò cột nó còn đối chiếu ngày↔tháng theo từng cost center
            # và soát cả sheet HQKD bản tháng. Các layout khác chưa có phần đối chiếu đó (mỗi bản
            # tháng một bố cục) -> chỉ dò cột, làm đúng phần chắc chắn đúng.
            th = thang_theo_cost_center(conn, source_file_thang) if source_file_thang else {}
            bc = doi_chieu(path, th)
        else:
            bc = soat_layout(path, layout)
            bc.setdefault("cot_bo_bat_thuong", [])
            bc.setdefault("doi_chieu_theo_cost_center", [])
            if not bc.get("so_cot_doc_duoc") and not bc.get("khong_ap_dung"):
                # 0 cột = hoặc sai layout, hoặc nguồn đổi bố cục. KHÔNG được im: im ở đây nghĩa là
                # bộ soát ngừng canh mà không ai biết — đúng thứ nó sinh ra để chặn.
                log(f"  soát cost center [{layout}]: ĐỌC ĐƯỢC 0 CỘT ở {os.path.basename(path)}"
                    " — nguồn có thể đã đổi bố cục, bộ soát đang KHÔNG canh đơn vị này")
    except Exception as ex:                              # noqa: BLE001 - xem docstring
        log(f"  soát cost center: BỎ QUA vì lỗi đọc ({type(ex).__name__}: {str(ex)[:120]})")
        return {}

    # Gộp cột chờ duyệt của CẢ HAI bản (ngày + tháng) vào một hàng đợi: cùng một dự án, cùng một
    # mã cần duyệt, duyệt một lần là cả hai pipeline cùng thấy (bản đồ DB dùng chung).
    cho_duyet = [dict(c, nguon="ngay") for c in bc["cot_chua_co_ma"]]
    bc["soat_thang"] = {}
    if source_file_thang:
        pt = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(path))),
                          "baocaotaichinhrieng", source_file_thang.split("::", 1)[-1])
        if os.path.isfile(pt):
            try:
                st = soat_thang(pt)
                bc["soat_thang"] = st
                co = {D._nd(c["ten_cot"]) for c in cho_duyet}
                for c in st.get("cot_chua_co_ma") or []:
                    if D._nd(c["ten_cot"]) not in co:
                        cho_duyet.append(dict(c, so_ngay_co_so=None, nguon="thang"))
                if abs(st.get("lech_so_voi_tong") or 0) > _NGUONG and not st.get("cot_chua_co_ma"):
                    log(f"  CANH BAO bản THÁNG: Σ các cột đã có mã hụt cột 'Tổng dự án'"
                        f" {st['lech_so_voi_tong']:,.0f} đ mà không cột nào chưa có mã"
                        " — kiểm tay, có thể sheet đổi bố cục")
            except Exception as ex:                      # noqa: BLE001
                log(f"  soát bản tháng: BỎ QUA ({type(ex).__name__}: {str(ex)[:100]})")

    for c in cho_duyet:
        ten, chuan = c["ten_cot"], D._nd(c["ten_cot"])
        giai_thich = next((k for k in bc["doi_chieu_theo_cost_center"]
                           if k["loai"] == "THIEU_COST_CENTER"
                           and abs(k["lech_dong"] - c["tong_dong"]) <= _NGUONG), None)
        dm = _DANH_MUC_LAYOUT.get(layout, {})
        bang_chung = {"ten_cot": ten, "tong_dong": c["tong_dong"],
                      "so_ngay_co_so": c.get("so_ngay_co_so"), "period": period,
                      "file": os.path.basename(path), "nguon": c.get("nguon", "ngay"),
                      "giai_thich_lech": giai_thich}
        conn.execute(
            """INSERT INTO cost_center_map
                 (layout, ten_cot_chuan, ten_cot_goc, cost_center, trang_thai, bang_chung,
                  ten_hien_thi, khoi, cong_ty)
               VALUES (%s,%s,%s,%s,'cho_duyet',%s,%s,%s,%s)
               ON CONFLICT (layout, ten_cot_chuan) DO UPDATE
                 SET bang_chung=EXCLUDED.bang_chung, ten_hien_thi=EXCLUDED.ten_hien_thi,
                     khoi=EXCLUDED.khoi, cong_ty=EXCLUDED.cong_ty, updated_at=now()
                 WHERE cost_center_map.trang_thai='cho_duyet'""",
            (layout, chuan, ten, ma_de_xuat(ten, layout), json.dumps(bang_chung, ensure_ascii=False),
             f"{dm.get('tien_to_ten', '')}{ten}".strip(), dm.get("khoi"), dm.get("cong_ty")))

        # dedupe_key KHÔNG mang kỳ: cột mới sẽ xuất hiện lại ở MỌI kỳ về sau, mỗi kỳ một việc
        # mới là chuông không bao giờ tắt. Một cột = một việc, duyệt xong là xong.
        conn.execute(
            """INSERT INTO source_change_events
                 (category, dedupe_key, company, report_type, file_name, period,
                  what_changed, chi_tiet, status, apply_log,
                  detected_at, last_seen_at, updated_at)
               VALUES ('costcenter',%s,%s,'baocaohqkdngay',%s,%s,%s,%s,'pending','[]',
                       now(),now(),now())
               ON CONFLICT (dedupe_key) DO UPDATE
                 SET last_seen_at=now(), chi_tiet=EXCLUDED.chi_tiet""",
            (f"costcenter|{layout}|{chuan}", company, os.path.basename(path), period,
             f"Cột '{ten}' chưa có mã cost center — {c['tong_dong']:,.0f} đ đang không lên dashboard",
             json.dumps(bang_chung, ensure_ascii=False)))
        _sn = c.get("so_ngay_co_so")
        log(f"  CHỜ DUYỆT [{c.get('nguon', 'ngay')}]: cột '{ten}' ({c['tong_dong']:,.0f} đ"
            + (f", {_sn} ngày có số" if _sn is not None else "") + f") -> đề xuất mã {ma_de_xuat(ten, layout)}")

    for k in bc["doi_chieu_theo_cost_center"]:
        if k["loai"] != "THIEU_COST_CENTER":
            log(f"  LỆCH [{k['cost_center']}] {k['lech_dong']:,.0f} đ — {k['vi_sao']}"
                f" (việc của {k['ai_xu_ly']})")
    for b in bc["cot_bo_bat_thuong"]:
        if b["loai"] == "so_rieng":
            log(f"  CANH BAO: cột '{b['ten_cot']}' đang bị bỏ nhưng mang số riêng"
                f" {b['tong_dong']:,.0f} đ — kiểm xem có phải tiền thật không")
    return bc


# ────────────────────────── SOÁT CỘT CỦA BẢN THÁNG ──────────────────────────
# Vì sao phải soát RIÊNG bản tháng: hai pipeline có HAI bản đồ khác nhau —
# `derive_hqkd_ngay._CC_DUAN` (ngày) và `agent_cli._DA_PROJECT_CC` (tháng). Lệch hai danh sách là
# dự án có số ở bản này mà mất trắng ở bản kia, IM LẶNG. Đã xảy ra: bảng THÁNG thiếu "binh phuoc"
# nên kỳ 2026-08 mất Bình Phước (DT 621.080.352, LNTT -126.050.367) khỏi DB, không ai biết cho tới
# lượt đối chiếu tay 18/09/2026. Bộ soát bản NGÀY ở trên KHÔNG bắt được ca đó.
#
# KHÁC bản ngày một điểm quan trọng: sheet HQKD bản THÁNG có cột "Tổng dự án" ĐÁNG TIN
# (T8/2026: 15.440.614.295 = đúng Σ 9 cột dự án), nên ở đây cột Tổng dùng làm chốt được — trong
# khi cột Tổng của file NGÀY sai 879 triệu (xem docstring đầu file). Đừng bê quy ước của bên này
# sang bên kia.

_SHEET_THANG = "HQKD"
_DONG_DT_THANG = "tong doanh thu rong"
_COT_BO_THANG = ("tong du an", "%dt")


def doc_cot_thang(path: str) -> dict:
    """Sheet HQKD bản THÁNG -> {cot: {...}, tong_cua_file: x}. {} nếu không đúng layout."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if _SHEET_THANG not in wb.sheetnames:
        wb.close()
        return {}
    rows = list(wb[_SHEET_THANG].iter_rows(min_row=1, max_row=40, max_col=40, values_only=True))
    wb.close()
    hdr_i = next((i for i, r in enumerate(rows)
                  if any(D._nd(c) == "tong du an" for c in r if c is not None)), None)
    dt = next((r for r in rows if D._nd(r[1] if len(r) > 1 else "") == _DONG_DT_THANG), None)
    if hdr_i is None or dt is None:
        return {}
    cot, tong = {}, 0.0
    for j, c in enumerate(rows[hdr_i]):
        if not isinstance(c, str) or not c.strip() or j >= len(dt):
            continue
        ten, n = c.strip(), D._nd(c)
        if n == "tong du an":
            tong = _f(dt[j])
            continue
        if n in _COT_BO_THANG:
            continue
        cot[ten] = {"ma": _ma_cua(ten), "tong": _f(dt[j])}
    return {"cot": cot, "tong_cua_file": tong}


def soat_thang(path: str) -> dict:
    """Cột dự án nào ở bản THÁNG chưa có mã + cột Tổng của file có khớp Σ các cột không."""
    q = doc_cot_thang(path)
    if not q:
        return {"bo_qua": "không đọc được sheet HQKD bản tháng"}
    cot = q["cot"]
    chua_co_ma = [{"ten_cot": t, "tong_dong": e["tong"]}
                  for t, e in cot.items() if not e["ma"] and abs(e["tong"]) > _NGUONG]
    da_co_ma = sum(e["tong"] for e in cot.values() if e["ma"])
    return {"cot_chua_co_ma": chua_co_ma,
            "tong_cua_file": q["tong_cua_file"],
            "tong_cac_cot_da_co_ma": da_co_ma,
            # Chốt tự kiểm: cột Tổng của bản tháng đáng tin, nên lệch ở đây = có cột đang rơi.
            "lech_so_voi_tong": q["tong_cua_file"] - da_co_ma,
            "can_admin_duyet": bool(chua_co_ma)}


# ══════════════════════ MỞ RỘNG SANG CÁC ĐƠN VỊ KHÁC (18/09/2026) ══════════════════════
# Chỉ 4 layout có CỘT COST CENTER trong file ngày nên chỉ 4 layout này có rủi ro "cột mới rơi im
# lặng": duan · srvf · xdv · antaxi. Các đơn vị còn lại (XVP, HTX×2 layout "kqkd"; GLOBALAI/TRAMSAC
# "tcode"; HO, HUNGTHINH, ANKHACHSAN) chỉ có MỘT cột giá trị, không tách cost center — `_kqkd_scan`
# rơi về cột "Tổng cộng" với cc=None. KHÔNG khai chúng ở đây: không có cột để mà thiếu, bật lên chỉ
# đẻ báo động giả.
#
# RỦI RO LỚN NHẤT CỦA VIỆC TỔNG QUÁT HOÁ LÀ BÁO ĐỘNG GIẢ, không phải bỏ sót. Một việc "chờ duyệt"
# sai làm admin mất niềm tin vào cả cái chuông, và tệ hơn hẳn việc chưa có gì. Vì vậy mỗi layout
# phải khai TƯỜNG MINH `bo_cot` — các cột KHÔNG PHẢI cost center nhưng vẫn nằm trên dòng header:
#   · srvf   "Mục/Khối/Mã số/Chỉ tiêu/Tổng/% DT" + cột pháp nhân "CHI NHÁNH VINFAST HÀ NỘI…"
#   · xdv    "Mã số/Chỉ tiêu/Kỳ này/TK nợ/TK có/Mã phí/Mã NS/Công thức"
#   · antaxi "TT/Chỉ tiêu/Tỉ lệ/Tổng cộng/Lũy kế…" + **"Xe thương quyền"** — cột SỐ LIỆU THẬT
#            nhưng là dòng kinh doanh, KHÔNG phải depot; không loại là mỗi lượt chạy đẻ một việc ma.
#
# `khop`: "chua" = tên cột CHỨA từ khoá (srvf/xdv/duan — tên cột có tiền tố "Showroom "/"XDV ");
#         "bang" = khớp CHÍNH XÁC (antaxi — `_kqkd_scan` dùng `==`). PHẢI GIỐNG DERIVER, lệch một
#         nhịp là bộ soát nói cột đã có mã trong khi deriver vẫn bỏ nó (hoặc ngược lại).
_LAYOUT_SOAT = {
    "duan": {"hdr": "tong du an", "khop": "chua", "map": lambda: D._CC_DUAN,
             "bo_cot": ("tong du an", "ho du an"),
             "dt": ("nhan", "doanh thu thuan"), "nhan_cot": "chi tieu"},
    "srvf": {"hdr": "ma so", "khop": "chua", "map": lambda: D._CC_SRVF,
             "bo_cot": ("muc", "khoi", "ma so", "chi tieu", "tong", "% dt", "chi nhanh"),
             "dt": ("ma", "A100"), "nhan_cot": "ma so"},
    "xdv": {"hdr": "ma so", "khop": "chua", "map": lambda: D._XDV_CC,
            "bo_cot": ("ma so", "chi tieu", "ky nay", "tk no", "tk co", "ma phi", "ma ns",
                       "cong thuc"),
            "dt": ("ma", "B100"), "nhan_cot": "ma so"},
    "antaxi": {"hdr": "chi tieu", "khop": "bang", "map": lambda: D._CC_ANTAXI,
               "bo_cot": ("tt", "chi tieu", "ti le", "tong cong", "luy ke", "xe thuong quyen"),
               "dt": ("nhan", "doanh thu thuan"), "nhan_cot": "chi tieu"},
}


def _la_sheet_ngay(ten_sheet: str) -> bool:
    """Sheet "1".."31" hoặc "01.09" (ngày.tháng). PHẢI CHẶN 1..31, không chỉ `isdigit()`.

    Bản đầu dùng `isdigit()` trần và nuốt luôn sheet **"131"/"331"** — đó là SỔ CÔNG NỢ theo tài
    khoản 131/331 trong file bản-chụp-một-ngày của An Taxi, không phải ngày nào cả. Hậu quả: 16
    file bị coi là "có sheet ngày mà đọc ra 0 cột" -> 16 dòng cảnh báo "đọc hỏng" mỗi lượt chạy,
    toàn bộ là giả.
    """
    s = str(ten_sheet or "").strip()
    if s.isdigit():
        return 1 <= int(s) <= 31
    p = s.split(".")
    return (len(p) == 2 and all(x.isdigit() for x in p)
            and 1 <= int(p[0]) <= 31 and 1 <= int(p[1]) <= 12)


def _ma_theo_layout(ten: str, layout: str):
    """Tên cột -> mã, theo ĐÚNG cách khớp của layout đó. Luôn kèm bản đồ DB admin đã duyệt."""
    cf = _LAYOUT_SOAT[layout]
    n = D._nd(ten)
    cap = list(cf["map"]()) + list(D._cc_da_duyet(layout))
    if cf["khop"] == "bang":
        return next((cc for kw, cc in cap if n == kw), None)
    return next((cc for kw, cc in cap if kw and kw in n), None)


def doc_cot_layout(path: str, layout: str) -> dict:
    """Quét mọi sheet ngày của MỘT layout -> {ten_cot: {ma, tong, so_ngay}}.

    Giá trị lấy ở ĐÚNG dòng doanh thu mà deriver dùng (`dt`), không phải cộng bừa mọi ô trong cột:
    một cột P&L trộn doanh thu, chi phí, lợi nhuận và tỉ lệ — cộng hết ra con số vô nghĩa, mà đây
    lại là con số admin nhìn để quyết có duyệt hay không.
    """
    cf = _LAYOUT_SOAT[layout]
    kieu_dt, khoa_dt = cf["dt"]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    cot, so_sheet_ngay = {}, 0
    for sn in wb.sheetnames:
        if not _la_sheet_ngay(sn):
            continue
        so_sheet_ngay += 1
        rows = list(wb[sn].iter_rows(min_row=1, max_row=200, max_col=40, values_only=True))
        hdr_i = next((i for i, r in enumerate(rows[:16])
                      if any(D._nd(c) == cf["hdr"] for c in r if c is not None)), None)
        if hdr_i is None:
            continue
        hdr = rows[hdr_i]
        # Cột nhãn chỉ tiêu có thể nằm ở DÒNG KHÁC với dòng tên cost center — layout "duan" là
        # đúng ca đó ("STT|CHỈ TIÊU|DỰ ÁN" ở một dòng, tên từng dự án ở dòng sau). Tìm trên CẢ
        # vùng header thay vì chỉ dòng `hdr`, nếu không duan đọc ra 0 cột và bộ soát im lặng
        # không làm gì — đúng kiểu hỏng nó sinh ra để chặn (bắt được lúc nghiệm thu 18/09).
        key_j = next((j for r in rows[:hdr_i + 3] for j, c in enumerate(r)
                      if D._nd(c) == cf["nhan_cot"]), None)
        if key_j is None:
            continue
        dong = None
        for r in rows[hdr_i:]:
            if key_j >= len(r) or r[key_j] in (None, ""):
                continue
            k = str(r[key_j]).strip()
            if (k.upper() == khoa_dt) if kieu_dt == "ma" else (D._nd(k) == khoa_dt):
                dong = r
                break
        if dong is None:
            continue
        for j, c in enumerate(hdr):
            if not isinstance(c, str) or not c.strip() or j >= len(dong):
                continue
            ten, n = c.strip(), D._nd(c)
            if any(b in n for b in cf["bo_cot"]):
                continue
            e = cot.setdefault(ten, {"ma": _ma_theo_layout(ten, layout), "tong": 0.0, "so_ngay": 0})
            v = _f(dong[j])
            e["tong"] += v
            e["so_ngay"] += 1 if abs(v) > _NGUONG else 0
    wb.close()
    return {"cot": cot, "so_sheet_ngay": so_sheet_ngay}


def soat_layout(path: str, layout: str) -> dict:
    """Cột cost center nào của layout này chưa có mã. Dùng cho MỌI layout kể cả 'duan'.

    `khong_ap_dung` tách bạch "file này KHÔNG có lưới cost center theo ngày" khỏi "có mà đọc
    hỏng" — An Taxi có 16 file bản-chụp-một-ngày (`.D.YYYYMMDD.`, layout `antaxi_bcqt`) đi qua
    cùng cron; coi chúng là bất thường thì mỗi lượt chạy in 16 dòng cảnh báo vô nghĩa, và cảnh
    báo nào kêu suốt thì chẳng ai đọc nữa.
    """
    q = doc_cot_layout(path, layout)
    cot = q["cot"]
    chua = [{"ten_cot": t, "tong_dong": e["tong"], "so_ngay_co_so": e["so_ngay"]}
            for t, e in cot.items() if not e["ma"] and abs(e["tong"]) > _NGUONG]
    return {"so_cot_doc_duoc": len(cot), "so_sheet_ngay": q["so_sheet_ngay"],
            "khong_ap_dung": q["so_sheet_ngay"] == 0,
            "cot_chua_co_ma": chua, "can_admin_duyet": bool(chua)}

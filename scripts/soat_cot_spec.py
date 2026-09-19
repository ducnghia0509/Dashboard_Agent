# -*- coding: utf-8 -*-
"""SOÁT CỘT COST CENTER cho nguồn TỰ ĐỘNG chạy bằng SPEC JSON (TEST_SR / TEST_XDV / …).

VÌ SAO CÓ FILE NÀY (19/09/2026). `soat_cost_center.py` chỉ canh các file `baocaohqkdngay` đọc
bằng LAYOUT viết tay (duan/srvf/xdv/antaxi). Nhưng số từ tháng 9 của XDV và Showroom đến từ nguồn
tự động Cyber, đọc bằng **spec JSON** — pipeline khác hẳn, cron kéo cũng khác. Hệ quả đo được:
cột **"Vinfast Xuân Mai"** xuất hiện trong file XDV, KHÔNG có trong `xdv_pnl_ngay.json`, mang
773.017 đ, rơi im lặng; người dùng phát hiện bằng mắt vì màn Tổng quan và QTVH lệch nhau ở đúng
3 ngày 14/15/16-09, còn bộ soát cũ vẫn báo "0 cột chưa có mã".

HAI CHỐT, CỐ Ý LÀM CẢ HAI:
  1. **So TÊN CỘT**: cột nào trong file mà không spec nào khai.
  2. **So SỐ**: `dòng tổng khối` vs `Σ các cột đã khai`. Chốt này MẠNH HƠN vì nó bắt được tiền
     đang rơi BẤT KỂ cột mới tên gì (kể cả khi nguồn đổi tên cột cũ) — và nếu có sẵn thì đã kêu
     ngay 14/09 thay vì để lệch 3 ngày mới có người thấy.

HAI BẪY BÁO ĐỘNG GIẢ ĐÃ DÍNH KHI DỰNG:
  · **Một file phục vụ NHIỀU KHỐI**: file `Baocaotaichinhrieng-HQKD` có CẢ 10 cột Showroom LẪN 14
    cột XDV (36 cột). Soát theo từng spec riêng lẻ thì cột XDV bị báo "thiếu" ở spec SR và ngược
    lại -> 17 báo động giả mỗi lượt. Phải gom UNION cột khai của MỌI spec cùng `nguon.folder`.
  · **Cột kỹ thuật**: Mã số/Chỉ tiêu/Lũy kế/Kỳ này/TK nợ/TK có/Mã phí/Mã NS/Công thức — không
    phải cost center, loại thẳng.
"""
import glob
import json
import os
import re
from collections import defaultdict

import openpyxl

RECEIVED = "/home/itadmin/AI_Dashboard_QT/Connect_VPS/received_reports"
SPEC_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "extract_specs")
_KY_THUAT = {"mã số", "chỉ tiêu", "lũy kế", "kỳ này", "tk nợ", "tk có", "mã phí", "mã ns",
             "công thức", "stt", "ghi chú", "thuyết minh", "đvt"}
_NGUONG = 1.0


def _headers_khai(spec) -> set:
    ra = set()
    for c in spec.get("cot_gia_tri") or []:
        h = c.get("header")
        for x in ([h] if isinstance(h, str) else (h or [])):
            ra.add(str(x).strip())
    return ra


def ban_do_folder() -> dict:
    """{folder: {"khai": set(header), "cot_tong": header dòng tổng khối, "specs": [...]}}.

    CHỈ nhận folder có ÍT NHẤT một spec khai `cost_center` — folder không tách cost center thì
    không có gì để mà thiếu, soát vào chỉ đẻ báo động giả.
    """
    ra = defaultdict(lambda: {"khai": set(), "cot_tong": None, "specs": []})
    for f in sorted(glob.glob(os.path.join(SPEC_DIR, "*.json"))):
        try:
            sp = json.load(open(f, encoding="utf-8"))
        except (OSError, ValueError):
            continue
        cg = sp.get("cot_gia_tri") or []
        if not any(isinstance(c, dict) and c.get("cost_center") for c in cg):
            continue
        fo = (sp.get("nguon") or {}).get("folder") or ""
        e = ra[fo]
        e["khai"] |= _headers_khai(sp)
        e["specs"].append(os.path.basename(f))
        # Cột tổng khối = mục cot_gia_tri KHÔNG có cost_center (vd "Kỳ này", dim2="Khối").
        for c in cg:
            if isinstance(c, dict) and not c.get("cost_center") and isinstance(c.get("header"), str):
                e["cot_tong"] = e["cot_tong"] or c["header"].strip()
    return dict(ra)


def _doc_dong(path: str, ma: str):
    """{header: giá trị} của dòng mã `ma` trong file. {} nếu không đọc được."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        rows = list(wb[wb.sheetnames[0]].iter_rows(min_row=1, max_row=80, max_col=80,
                                                   values_only=True))
    finally:
        wb.close()
    hi = next((i for i, r in enumerate(rows)
               if any(isinstance(c, str) and c.strip() == "Mã số" for c in r)), None)
    if hi is None:
        return {}
    hdr = [str(c).strip() if isinstance(c, str) else "" for c in rows[hi]]
    mj = hdr.index("Mã số")
    r = next((x for x in rows[hi + 1:]
              if mj < len(x) and str(x[mj]).strip() == ma), None)
    if r is None:
        return {}
    return {h: (float(r[j]) if j < len(r) and isinstance(r[j], (int, float)) else 0.0)
            for j, h in enumerate(hdr) if h}


# Mã dòng DOANH THU TỔNG, thử theo thứ tự: XDV dùng B-series, Showroom dùng A-series. Dò thay vì
# bắt khai trong spec — spec không có chỗ nào nói "dòng nào là doanh thu", và đoán sai thì chốt
# so-số im lặng không chạy (còn tệ hơn báo sai).
_MA_DT = ("B100", "A100")
# Cột tổng khối mặc định khi spec KHÔNG khai mục nào thiếu `cost_center` (spec SR khai đủ 10 cột
# showroom, không có mục khối) — mọi file họ này đều có cột "Kỳ này" là số của cả khối.
_COT_TONG_MAC_DINH = "Kỳ này"


def soat_folder(folder: str, cf: dict, ma_dt: str = None) -> dict:
    """Soát thư mục nguồn: cột chưa khai + chốt tổng-khối vs Σ-cột-khai, trên 2 file mới nhất.

    Dùng HAI file liền kề vì các nguồn này là ẢNH CHỤP LUỸ KẾ (`tru_ngay_truoc`): giá trị của
    riêng một ngày = hiệu hai bản. So trên bản luỹ kế sẽ trộn cả tháng vào một con số.
    """
    fs = [p for p in sorted(glob.glob(os.path.join(RECEIVED, folder, "*.xlsx")))
          if os.path.getsize(p) > 5000]
    if len(fs) < 2:
        return {"bo_qua": "chưa đủ 2 file để tính số phát sinh của một ngày"}
    ma_list = [ma_dt] if ma_dt else list(_MA_DT)
    sau = truoc = {}
    for m in ma_list:
        sau = _doc_dong(fs[-1], m)
        if sau:
            truoc, ma_dt = _doc_dong(fs[-2], m), m
            break
    if not sau:
        return {"bo_qua": f"không thấy dòng {'/'.join(ma_list)} ở {os.path.basename(fs[-1])}"}
    d = lambda h: sau.get(h, 0.0) - truoc.get(h, 0.0)            # noqa: E731
    la = [h for h in sau
          if h.strip().lower() not in _KY_THUAT and h.strip() not in cf["khai"]]
    chua_khai = [{"ten_cot": h, "phat_sinh": d(h), "luy_ke": sau.get(h, 0.0)}
                 for h in la if abs(d(h)) > _NGUONG or abs(sau.get(h, 0.0)) > _NGUONG]
    cot_tong = cf["cot_tong"] or (_COT_TONG_MAC_DINH if _COT_TONG_MAC_DINH in sau else None)
    tong_khoi = d(cot_tong) if cot_tong else None
    sum_khai = sum(d(h) for h in cf["khai"] if h != cot_tong and h in sau)
    return {"file": os.path.basename(fs[-1]), "ma_dt": ma_dt, "cot_tong": cot_tong,
            "so_cot_file": len(sau),
            "cot_chua_khai": chua_khai,
            "tong_khoi": tong_khoi, "sum_cot_khai": sum_khai,
            "lech_tong": (tong_khoi - sum_khai) if tong_khoi is not None else None,
            "can_admin_duyet": bool(chua_khai)}


# ────────────────────── XẾP VIỆC CHO ADMIN DUYỆT ──────────────────────
# Dùng CHUNG hàng đợi với bộ soát layout (`soat_cost_center.xep_viec`): chuông 🔔
# `source_change_events` + bảng `cost_center_map`. Khác đúng một chỗ — `layout` ở đây là
# `spec:<nguon.folder>` chứ không phải tên layout, để `spec_extract.load_spec()` tra đúng bản đồ.

import json as _json                                             # noqa: E402

_KHOI_THEO_FOLDER = {
    "TEST_XDV/baocaotaichinhrienghqkd": ("Khối KD Vinfast - XDV", "TC", "_XDV"),
    "XDV/baocaotaichinhrieng":          ("Khối KD Vinfast - XDV", "TC", "_XDV"),
    "TEST_SR/baocaotaichinhrienghqkd":  ("Khối KD Vinfast - Showroom", "TC", "_SR"),
}


def _ma_de_xuat(ten: str, hau_to: str) -> str:
    n = re.sub(r"^(showroom|xdv|vf|vinfast|chi nhanh vinfast)\s+", "",
               _bo_dau(ten), flags=re.I).strip()
    n = "".join(c for c in n.upper() if c.isalnum())
    return f"{n[:24]}{hau_to}" if n else ""


def _bo_dau(s: str) -> str:
    import unicodedata
    t = str(s or "").strip().lower().replace("đ", "d")
    t = "".join(c for c in unicodedata.normalize("NFD", t)
                if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t)


def xep_viec(conn, folder: str, cf: dict, kq: dict, log=print) -> None:
    """Ghi cột chưa khai thành việc CHỜ DUYỆT. KHÔNG ném lỗi ra ngoài (lớp giám sát)."""
    khoi, cong_ty, hau_to = _KHOI_THEO_FOLDER.get(folder, (None, None, "_CC"))
    for c in kq.get("cot_chua_khai") or []:
        ten, chuan = c["ten_cot"], _bo_dau(c["ten_cot"])
        bc = {"ten_cot": ten, "tong_dong": c["luy_ke"], "phat_sinh_ngay": c["phat_sinh"],
              "so_ngay_co_so": None, "file": kq.get("file"), "nguon": "spec",
              "folder": folder, "lech_tong_khoi": kq.get("lech_tong"),
              "giai_thich_lech": ({"vi_sao": "khớp khít chỗ lệch giữa dòng tổng khối và Σ cột đã khai"}
                                  if abs((kq.get("lech_tong") or 0) - c["phat_sinh"]) <= _NGUONG
                                  else None)}
        conn.execute(
            """INSERT INTO cost_center_map
                 (layout, ten_cot_chuan, ten_cot_goc, cost_center, trang_thai, bang_chung,
                  ten_hien_thi, khoi, cong_ty)
               VALUES (%s,%s,%s,%s,'cho_duyet',%s,%s,%s,%s)
               ON CONFLICT (layout, ten_cot_chuan) DO UPDATE
                 SET bang_chung=EXCLUDED.bang_chung, updated_at=now()
                 WHERE cost_center_map.trang_thai='cho_duyet'""",
            (f"spec:{folder}", chuan, ten, _ma_de_xuat(ten, hau_to),
             _json.dumps(bc, ensure_ascii=False), ten, khoi, cong_ty))
        conn.execute(
            """INSERT INTO source_change_events
                 (category, dedupe_key, company, report_type, file_name,
                  what_changed, chi_tiet, status, apply_log, detected_at, last_seen_at, updated_at)
               VALUES ('costcenter',%s,%s,'baocaotaichinhrienghqkd',%s,%s,%s,'pending','[]',
                       now(),now(),now())
               ON CONFLICT (dedupe_key) DO UPDATE
                 SET last_seen_at=now(), chi_tiet=EXCLUDED.chi_tiet""",
            (f"costcenter|spec:{folder}|{chuan}", folder.split("/")[0], kq.get("file"),
             f"Cột '{ten}' chưa khai trong spec — {c['luy_ke']:,.0f} đ đang không lên dashboard",
             _json.dumps(bc, ensure_ascii=False)))
        log(f"  CHỜ DUYỆT [spec {folder.split('/')[0]}]: cột '{ten}'"
            f" (luỹ kế {c['luy_ke']:,.0f} đ) -> đề xuất mã {_ma_de_xuat(ten, hau_to)}")


def soat_va_xep(conn, log=print) -> dict:
    """Quét MỌI thư mục nguồn chạy bằng spec có cost center rồi xếp việc. -> {folder: kết quả}."""
    ra = {}
    for folder, cf in ban_do_folder().items():
        try:
            kq = soat_folder(folder, cf)
            ra[folder] = kq
            if kq.get("bo_qua"):
                continue
            lech = kq.get("lech_tong")
            if lech is not None and abs(lech) > _NGUONG and not kq.get("cot_chua_khai"):
                # Có tiền rơi mà KHÔNG cột lạ nào -> nguồn đổi TÊN cột cũ, hoặc dòng tổng sai.
                # Phải kêu: đây đúng là ca mà chốt so-tên một mình sẽ bỏ lọt.
                log(f"  CANH BAO [spec {folder}]: dòng tổng khối lệch Σ cột đã khai"
                    f" {lech:,.0f} đ mà không cột nào lạ — kiểm tên cột có bị đổi không")
            xep_viec(conn, folder, cf, kq, log=log)
        except Exception as ex:                          # noqa: BLE001
            log(f"  soát cột spec [{folder}]: BỎ QUA ({type(ex).__name__}: {str(ex)[:120]})")
    return ra

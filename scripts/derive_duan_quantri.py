#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deriver: SỐ QUẢN TRỊ khối Dự án — các file báo cáo dự án của KTKH (mapping 09/10/2026).

Mapping `Tài liệu/4.DU_AN.Mapping_Dashboard copy.xlsx`, sheet "2. Mapping 09.10 - Final". Máy gửi
bắt đầu chào nhánh `/data/4.DU_AN` ngày 09/10/2026, mỗi họ file một `report_type` riêng:

  duanbctonghopnam  `B.4.TC.TH_DA.N.2026.BCTONGHOP.xlsx` — mỗi dự án một sheet, dòng = chỉ tiêu,
                    cột "Tháng 1".."Tháng 12" + vài cột LUỸ KẾ ("Tổng hợp", "Đã NT", "Chưa NT").
  duanttbctonghop   `B.4.TC.TT_DA.TH.<...>.BCTONGHOPTANTHINH.xlsx` — CÙNG bố cục, chỉ sheet Tân
                    Thịnh, số Y HỆT sheet "Tân Thịnh" của file trên (đo 09/10: khớp từng ô).
  duanpqbcthang     `B.4.TC.PQ_DA.M.<yyyymm>.BCTHANG.xlsx` — hồ sơ nghiệm thu tháng Thổ Chu.
  duanttbcdashboard `BC DashBoard.xlsx` (Tân Thịnh) — chỉ lấy "TỔNG GIÁ TRỊ TRƯỚC THUẾ" sheet GTTT.
  duantheodoihopdong `Bang_theo_doi_hop_dong_xay_dung_*.xlsx` — công nợ theo hợp đồng, 1 sheet/dự án.

VÌ SAO LÀ DERIVER, KHÔNG PHẢI SPEC JSON: mapping trỏ vào Ô CỐ ĐỊNH (Y4, Z4+Z5+Z6, B9, B12…) trên
các sheet mỗi cái một bố cục, sheet Đảo Thổ Chu đi CẶP cột (khối lượng | thành tiền) tới tháng 8
rồi một cột từ tháng 9, và một nửa số là ảnh chụp "luỹ kế đến hiện tại" không mang kỳ. Engine spec
không có chế độ đọc ô lẻ. Ở đây vẫn KHÔNG đọc theo số ô: dòng dò theo NHÃN cột B, cột dò theo
TIÊU ĐỀ — ô mapping ghi chỉ dùng để đối chiếu (mapping đã lệch file thật ở 3 chỗ, xem `_bay`).

BA REPORT_TYPE (đơn vị tỷ, cong_ty TC, khối "Khối KD Dự án"):
  DUAN_QT_THANG  chỉ tiêu THÁNG × dự án, neo ngày cuối tháng. dim1 = chỉ tiêu chuẩn, dim2 = hợp
                 đồng (Cao Bằng tách doanh thu 3 HĐ), dim3 = dòng CHI TIẾT (vd "Dầu - Máy xúc KL"
                 dưới "Chi phí nhiên liệu"). Cộng chỉ tiêu = các dòng dim3 RỖNG; dòng có dim3 là
                 chi tiết của dòng cha, cộng vào là đếm đôi.
  DUAN_QT_HD     ẢNH CHỤP luỹ kế theo hợp đồng (giá trị HĐ trước thuế · sản lượng luỹ kế · đã/chưa
                 nghiệm thu), neo CUỐI THÁNG CUỐI CÙNG CÓ SỐ của sheet (luỹ kế tới tháng báo cáo).
  DUAN_CN_HD     ẢNH CHỤP công nợ theo hợp đồng (bảng theo dõi HĐ), neo NGÀY PHÁT SINH CUỐI (sheet
                 "Chi tiết phát sinh", không vượt ngày sửa file).
  DUAN_QT_NGAY   (vòng 2) chỉ tiêu NGÀY × dự án từ báo cáo ngày Thổ Chu (`duanpqbcngay`) và Cao
                 Bằng (`duancbbcngay`): nhân công, nhiên liệu, khấu hao, lương thợ lái, sửa chữa,
                 khác, thầu phụ, + 'Khối lượng (m³)' (payload.unit='m3', KHÔNG phải tỷ).
  DUAN_TP_THANG  (vòng 2) thầu phụ Thổ Chu theo tháng (`duanpqdbthauphu`): sản lượng · nghiệm thu
                 · dở dang.
  Hai cụm ảnh chụp đọc BẢN MỚI NHẤT ≤ cuối cửa sổ xem, không cộng qua kỳ. KHÔNG neo theo ngày sửa
  file: dữ liệu lọc theo dataset từng tháng, neo 08/10 thì xem tháng 8-9 không thấy số luỹ kế nào.

Chạy: .venv/bin/python scripts/derive_duan_quantri.py [--file X ...] [--write]
       (không --file: quét cả 5 thư mục trên dưới received_reports/DUAN)
"""
import argparse
import calendar
import datetime as dt
import glob
import json
import os
import re
import sys

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
import spec_extract as SE  # noqa: E402  — dùng chung _nd/_so/_cc_duan/_ghi/_source_id

RT_THANG, RT_HD, RT_CN = "DUAN_QT_THANG", "DUAN_QT_HD", "DUAN_CN_HD"
KHOI, CONG_TY, HE_SO = "Khối KD Dự án", "TC", 1e-9

_bay = [
    "MAPPING LỆCH FILE 1 — dòng 15 Tân Thịnh ghi 'H15 -> H15' nhưng dòng 15 là TỔNG CHI PHÍ; lợi "
    "nhuận nằm dòng 16 'LỢI NHUẬN'. Đọc theo nhãn nên tự đúng.",
    "MAPPING LỆCH FILE 2 — dòng 46-47 Cao Bằng ghi ô 'B33' (sheet chỉ có 20 dòng); nhãn mapping là "
    "'Chi phí thầu phụ' -> lấy dòng 'Chi phí thầu phụ' theo tháng.",
    "HAI FILE TRÙNG SỐ — sheet 'Tân Thịnh' của BCTONGHOP và file BCTONGHOPTANTHINH giống hệt nhau. "
    "Mapping trỏ Tân Thịnh về file riêng -> sheet 'Tân Thịnh' của BCTONGHOP BỊ BỎ, đọc cả hai là "
    "cộng đôi.",
    "CẶP CỘT Ở ĐẢO THỔ CHU — tiêu đề 'Tháng N' đứng ở cột KHỐI LƯỢNG, thành tiền ở cột trống ngay "
    "sau (F44 = lợi nhuận T2 như mapping ghi); từ tháng 9 mỗi tháng một cột. Luật: cột ngay sau "
    "tiêu đề trống -> giá trị nằm ở đó, không thì ở chính cột tiêu đề.",
    "LUỸ KẾ TÂN THỊNH TÍNH TỪ 2024 — 'Giá trị đã được NT' (V7) là luỹ kế từ khi ký HĐ (2024), "
    "trong khi Σ tháng 2026 chỉ là năm nay. Sản lượng luỹ kế vì vậy lấy cột 'TỔNG HỢP' (U7) cùng "
    "gốc với V7; Σ 2026 (mapping dòng 33) cho dở dang âm 39,7 tỷ — vô nghĩa. Thổ Chu/Cao Bằng bắt "
    "đầu 2026 nên hai cách cho cùng số.",
    "GIÁ TRỊ HĐ Ở BCTONGHOP LÀ TRƯỚC THUẾ — Thổ Chu 129,63 tỷ × 1,08 = 140 tỷ của bảng theo dõi HĐ; "
    "Cao Bằng 124,99 × 1,08 = 134,99. Hai nguồn không lệch, chỉ khác cơ sở thuế.",
]

# Nhãn dòng (đã `_nd`) -> chỉ tiêu chuẩn. Khớp BẰNG, không startswith: "doanhthu" là tiền tố của
# "doanhthukhac". Dòng không có ở đây và đứng SAU một dòng có ở đây = dòng chi tiết của dòng đó.
_CHI_TIEU = {
    "doanhthu": "Doanh thu sản xuất",
    "doanhthukhac": "Doanh thu khác",
    "doanhthukhacphithutucdt": "Doanh thu khác",
    "giamtrudoanhthu": "Giảm trừ doanh thu",
    "tongdoanhthudong": "Tổng doanh thu",
    "chiphinhienlieu": "Chi phí nhiên liệu",
    "chiphinhancongtructiep": "Chi phí nhân công trực tiếp",
    "chiphikhauhao": "Chi phí khấu hao",
    "chiphithauphu": "Chi phí thầu phụ",
    "chiphinhapvatlieu": "Chi phí vật liệu",
    "cpnvlhdps166168": "Chi phí vật liệu",
    "chiphisuachua": "Chi phí sửa chữa",
    "chiphibanbuon": "Chi phí bán buôn",
    "chiphikhactaiduan": "Chi phí khác tại dự án",
    "chiphikhaclaivayvacpql": "Chi phí khác (lãi vay, quản lý)",
    "chiphikeoxemaymoctb": "Chi phí kéo xe máy móc TB",
    "chiphitainanlaodong": "Chi phí tai nạn lao động",
    "tongchiphi": "Tổng chi phí",
    "loinhuan": "Lợi nhuận",
}
# Cột luỹ kế (tiêu đề đã `_nd`) -> chỉ tiêu ảnh chụp của DUAN_QT_HD.
_LUY_KE = {
    "tonghop": "Sản lượng luỹ kế", "dant": "Đã nghiệm thu luỹ kế", "dtdant": "Đã nghiệm thu luỹ kế",
    "giatridaduocnt": "Đã nghiệm thu luỹ kế",
    "chuant": "Chưa nghiệm thu", "dtchuant": "Chưa nghiệm thu", "giatrichuaduocnt": "Chưa nghiệm thu",
    "kehoach": "Giá trị hợp đồng", "kehoachnamtruocthue": "Giá trị hợp đồng",
}
_RE_THANG = re.compile(r"^thang(\d{1,2})$")


def _ngay_sua(path):
    """Ngày sửa file ở NGUỒN (sidecar của receiver) — trần cho ngày neo, và đường lùi cuối."""
    side = os.path.splitext(path)[0] + ".json"
    try:
        with open(side, encoding="utf-8") as fh:
            m = json.load(fh)
        s = m.get("modified_at") or m.get("saved_at")
        return s[:10] if s else None
    except (OSError, ValueError):
        return dt.date.fromtimestamp(os.path.getmtime(path)).isoformat()


def _cuoi_thang(y, m):
    return dt.date(y, m, calendar.monthrange(y, m)[1]).isoformat()


def _rec(ngay, cc, dim1, v, dim2=None, dim3=None, **payload):
    return {"ngay": ngay, "cong_ty": CONG_TY, "khoi": KHOI, "cost_center": cc,
            "amount": v * HE_SO, "dim1": dim1, "dim2": dim2, "dim3": dim3,
            "payload": {"unit": "ty", **payload}}


def _dong_tieu_de(rows):
    """Chỉ số dòng có >= 3 ô 'Tháng N' — dòng tiêu đề tháng (dòng 3 ở mọi sheet đo được)."""
    for i, r in enumerate(rows[:10]):
        if sum(1 for c in r if _RE_THANG.match(SE._nd(c))) >= 3:
            return i
    return None


def _cot_gia_tri(hdr, j):
    """Cột chứa SỐ TIỀN của tiêu đề ở cột j — xem `_bay` cặp cột Đảo Thổ Chu."""
    return j + 1 if j + 1 < len(hdr) and hdr[j + 1] in (None, "") else j


def _sheet_ma_tran(ws, nam, cc, nguon):
    """Một sheet dự án của BCTONGHOP -> (bản ghi tháng, {(dim1, dim2): giá trị luỹ kế})."""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    i_hdr = _dong_tieu_de(rows)
    if i_hdr is None:
        return [], {}, [f"sheet {ws.title}: không thấy dòng tiêu đề 'Tháng N'"]
    hdr = rows[i_hdr]
    thang = {}
    for j, c in enumerate(hdr):
        m = _RE_THANG.match(SE._nd(c))
        if m and 1 <= int(m.group(1)) <= 12:
            thang[int(m.group(1))] = _cot_gia_tri(hdr, j)
    luy_ke = {}
    for j, c in enumerate(hdr):
        k = _LUY_KE.get(SE._nd(c))
        if k and k not in {v for v, _ in luy_ke.values()}:
            luy_ke[j] = (k, _cot_gia_tri(hdr, j) if k == "Giá trị hợp đồng" else j)
    recs, snap, warn, cha = [], {}, [], None
    for r in rows[i_hdr + 1:]:
        nhan = r[1] if len(r) > 1 else None
        n = SE._nd(nhan)
        if not n:
            continue
        hop_dong = None
        if n.startswith("dthd"):                       # Cao Bằng: doanh thu theo từng HĐ
            ct, hop_dong = "Doanh thu sản xuất", str(nhan).strip()
        else:
            ct = _CHI_TIEU.get(n)
        chi_tiet = None
        if ct:
            cha = ct
        elif cha:
            ct, chi_tiet = cha, str(nhan).strip()
        else:
            warn.append(f"sheet {ws.title}: dòng lạ '{nhan}' không có dòng cha — bỏ")
            continue
        for mo, j in sorted(thang.items()):
            v = SE._so(r[j]) if j < len(r) else 0.0
            if v:
                recs.append(_rec(_cuoi_thang(nam, mo), cc, ct, v, hop_dong, chi_tiet,
                                 nguon=nguon, dong=str(nhan).strip()))
        if chi_tiet is None and ct in ("Doanh thu sản xuất", "Tổng doanh thu"):
            for _j, (k, jv) in luy_ke.items():
                # Ô luỹ kế bằng 0 VẪN ghi: "Chưa NT" = 0 (Cao Bằng) là số thật — dở dang bằng 0,
                # bỏ đi thì màn đọc thành "chưa có nguồn". Chỉ bỏ ô trống/chữ.
                o = r[jv] if jv < len(r) else None
                if isinstance(o, (int, float)) and not isinstance(o, bool):
                    snap[(ct, k, hop_dong)] = float(o)
    return recs, snap, warn


def _luy_ke_hd(snap, cc, ngay, nguon):
    """Ảnh chụp luỹ kế -> DUAN_QT_HD, lấy CẢ BỘ từ MỘT dòng để sản lượng − đã NT = chưa NT đúng
    đẳng thức của file: dòng doanh thu sản xuất nếu dòng đó có cột "đã nghiệm thu" (Thổ Chu X4/Y4,
    Cao Bằng Y4:Y6/Z4:Z6), không thì dòng "Tổng doanh thu" (Tân Thịnh U7/V7/W7 — dòng 4 chỉ có U4,
    trộn U4 với V7 là lệch 1,2 tỷ doanh thu khác). Giá trị HĐ luôn ở dòng doanh thu sản xuất."""
    dong = ("Doanh thu sản xuất" if any(ct == "Doanh thu sản xuất" and k == "Đã nghiệm thu luỹ kế"
                                        for ct, k, _h in snap) else "Tổng doanh thu")
    out = []
    for (ct, k, hd), v in sorted(snap.items(), key=lambda x: str(x[0])):
        if (k == "Giá trị hợp đồng" and ct == "Doanh thu sản xuất") or (k != "Giá trị hợp đồng"
                                                                       and ct == dong):
            out.append(_rec(ngay, cc, k, v, hd, None, nguon=nguon))
    return out


def doc_bctonghop(path, chi_sheet=None):
    m = re.search(r"\.(?:N|TH)\.(\d{4})", os.path.basename(path))
    nam = int(m.group(1)) if m else dt.date.today().year
    ngay = _ngay_sua(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    thang, hd, warn = [], [], []
    for s in wb.sheetnames:
        n = SE._nd(s)
        if chi_sheet and n != SE._nd(chi_sheet):
            continue
        if not chi_sheet and n == "tanthinh":           # xem `_bay` "HAI FILE TRÙNG SỐ"
            continue
        cc = SE._cc_duan(s)
        if not isinstance(cc, str):                     # 'TH', 'Sheet1' — sheet tổng, không phải dự án
            continue
        r, snap, w = _sheet_ma_tran(wb[s], nam, cc, os.path.basename(path))
        thang += r
        # Neo luỹ kế vào cuối tháng cuối có số của CHÍNH sheet đó (Quang Sơn dừng T6, Thổ Chu T8).
        moc = max((x["ngay"] for x in r), default=ngay)
        hd += _luy_ke_hd(snap, cc, moc, os.path.basename(path))
        warn += w
    return {RT_THANG: thang, RT_HD: hd}, warn


def doc_bcthang_pq(path):
    """Nghiệm thu tháng Thổ Chu: sheet "Tổng hợp- BC", dòng "DOANH THU BÁN HÀNG VÀ CUNG CẤP DỊCH
    VỤ", cột "Thành tiền" của nhóm "Khối lượng Thịnh Cường thực hiện" (ô I6 của mapping). Kỳ lấy
    từ TÊN FILE — cột A ghi "Tháng 01" ở mọi bản (chép lại, không cập nhật)."""
    m = re.search(r"\.M\.(\d{4})(\d{2})\.", os.path.basename(path))
    if not m:
        return {RT_THANG: []}, ["tên file không có kỳ .M.yyyymm."]
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ten = next((s for s in wb.sheetnames if SE._nd(s) == "tonghopbc"), None)
    if not ten:
        return {RT_THANG: []}, ["không có sheet 'Tổng hợp- BC'"]
    rows = [list(r) for r in wb[ten].iter_rows(max_row=40, values_only=True)]
    i_nhom = next((i for i, r in enumerate(rows) if any(SE._nd(c) == "khoiluongthinhcuongthuchien"
                                                        for c in r)), None)
    if i_nhom is None:
        return {RT_THANG: []}, ["không thấy nhóm cột 'Khối lượng Thịnh Cường thực hiện'"]
    g = next(j for j, c in enumerate(rows[i_nhom]) if SE._nd(c) == "khoiluongthinhcuongthuchien")
    g2 = next((j for j in range(g + 1, len(rows[i_nhom])) if rows[i_nhom][j] not in (None, "")),
              len(rows[i_nhom]))
    hdr = rows[i_nhom + 1]
    jv = next((j for j in range(g, g2) if j < len(hdr) and SE._nd(hdr[j]) == "thanhtien"), None)
    r = next((r for r in rows if any(SE._nd(c) == "doanhthubanhangvacungcapdichvu" for c in r)), None)
    if jv is None or r is None:
        return {RT_THANG: []}, ["không thấy cột 'Thành tiền' hoặc dòng doanh thu bán hàng"]
    v = SE._so(r[jv])
    ngay = _cuoi_thang(int(m.group(1)), int(m.group(2)))
    return {RT_THANG: [_rec(ngay, "TC_DA", "Doanh thu nghiệm thu", v, nguon=os.path.basename(path))]
            if v else []}, []


def doc_bcdashboard_tt(path):
    """Giá trị HĐ Tân Thịnh trước thuế: sheet GTTT*, dòng "TỔNG GIÁ TRỊ TRƯỚC THUẾ" (không phải
    dòng "... LÀM TRÒN"), cột I của mapping — kiểm cột I là số."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ten = next((s for s in wb.sheetnames if SE._nd(s).startswith("gttt")), None)
    if not ten:
        return {RT_HD: []}, ["không có sheet GTTT"]
    # Neo vào "Ngày xem BC" (sheet "BC ngày", dòng 2) — ngày chốt của chính báo cáo này.
    moc = _ngay_sua(path)
    if "BC ngày" in wb.sheetnames:
        dong2 = list(next(wb["BC ngày"].iter_rows(min_row=2, max_row=2, values_only=True), ()))
        j = next((j for j, c in enumerate(dong2) if SE._nd(c) == "ngayxembc"), None)
        if j is not None and j + 1 < len(dong2) and isinstance(dong2[j + 1], dt.datetime):
            moc = min(moc, dong2[j + 1].date().isoformat())
    for r in wb[ten].iter_rows(values_only=True):
        if len(r) > 8 and SE._nd(r[1]) == "tonggiatritruocthue" and isinstance(r[8], (int, float)):
            return {RT_HD: [_rec(moc, "TT_DA", "Giá trị hợp đồng", float(r[8]),
                                 nguon=os.path.basename(path), sheet=ten)]}, []
    return {RT_HD: []}, ["không thấy dòng 'TỔNG GIÁ TRỊ TRƯỚC THUẾ' có số ở cột I"]


def _ngay_o(v):
    if isinstance(v, dt.datetime):
        return v.date()
    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})", str(v or ""))
    return dt.date(int(m[3]), int(m[2]), int(m[1])) if m else None


def _moc_phat_sinh(wb, tran):
    """Ngày phát sinh LỚN NHẤT không vượt ngày sửa file (sheet "Chi tiết phát sinh", cột B).

    Đo 09/10/2026: có 2 dòng ghi 02/11/2026 — ngày tương lai, nhiều khả năng gõ đảo ngày/tháng.
    Lấy MAX trần là neo cả bảng vào tháng 11 (kỳ chưa tới, không khai sinh được -> mất trắng);
    nên chặn trần và BÁO RA các dòng đó để kế toán sửa."""
    ten = next((s for s in wb.sheetnames if SE._nd(s) == "chitietphatsinh"), None)
    if not ten:
        return tran, ["không có sheet 'Chi tiết phát sinh' — neo ngày sửa file"]
    tran_d = dt.date.fromisoformat(tran)
    ngay = [d for r in wb[ten].iter_rows(min_row=4, values_only=True)
            for d in [_ngay_o(r[1] if len(r) > 1 else None)] if d]
    tuong_lai = sorted({d.isoformat() for d in ngay if d > tran_d})
    hop_le = [d for d in ngay if d <= tran_d]
    warn = ([f"'Chi tiết phát sinh' có ngày sau ngày sửa file: {', '.join(tuong_lai)}"]
            if tuong_lai else [])
    return (max(hop_le).isoformat() if hop_le else tran), warn


def doc_theodoi_hopdong(path):
    """Bảng theo dõi HĐ: mỗi sheet dự án có khối "THÔNG TIN HỢP ĐỒNG" nhãn cột A, số cột B (B9..B19
    của mapping). Dò theo nhãn, bỏ hậu tố '(VNĐ)'. Sheet tổng/hướng dẫn/chi tiết phát sinh bỏ."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ngay, warn = _moc_phat_sinh(wb, _ngay_sua(path))
    out = []
    for s in wb.sheetnames:
        cc = SE._cc_duan(s)
        if not isinstance(cc, str):
            continue
        info, so = {}, []
        for r in wb[s].iter_rows(max_row=21, max_col=2, values_only=True):
            nhan, gt = (r + (None, None))[:2]
            if not nhan:
                continue
            ten = re.sub(r"\s*\(VNĐ\)\s*$", "", str(nhan).strip(), flags=re.I)
            if isinstance(gt, (int, float)) and "(%)" not in ten:
                so.append((ten, float(gt)))
            elif gt not in (None, ""):
                info[SE._nd(ten)] = gt
        hd = str(info.get("sohopdong") or "").strip() or None
        for ten, gt in so:
            out.append(_rec(ngay, cc, ten, gt, hd, None, nguon=os.path.basename(path),
                            chu_dau_tu=str(info.get("chudautudoitac") or "").strip() or None,
                            trang_thai=str(info.get("trangthaihopdong") or "").strip() or None))
        if not so:
            warn.append(f"sheet {s}: không có dòng số nào trong khối thông tin HĐ")
    return {RT_CN: out}, warn


# ───────────────────── vòng 2 · báo cáo NGÀY Thổ Chu / Cao Bằng · thầu phụ Thổ Chu ─────────────────────
# Mapping dòng 66-77 (Thổ Chu — sheet "TH báo cáo ngày"): chỉ tiêu = Σ các dòng nhãn cột B. Nhãn đã
# `_nd`. Dòng "Thợ sửa chữa + bảo dưỡng" xuất hiện HAI lần (24 và 26) — mapping chỉ lấy dòng 24 nên
# chỉ khớp LẦN ĐẦU của mỗi nhãn.
_PQ_NGAY = {
    "Chi phí nhân công": ["chiphinhancong"],                                   # dòng 17
    "Chi phí nhiên liệu": ["chiphinhienlieu"],                                 # dòng 7
    "Chi phí khấu hao": ["chiphikhauhao"],                                     # dòng 43
    "Lương thợ lái": ["laimayxuckl", "laixekhoiluong", "laimayui"],            # dòng 19+20+21
    "Chi phí sửa chữa": ["thosuachuabaoduong", "chiphisuachua"],               # dòng 24+30
    "Chi phí khác": ["cuocvanchuyenvattuthietbi", "chiquytienmat"],            # dòng 41+42
    # Thiết bị xúc / m³ (mapping dòng 74-77): tử số theo m³ khối lượng (dòng 4).
    "Dầu máy xúc": ["daumayxuckl"],                                            # dòng 8
    "Lương lái máy xúc": ["laimayxuckl"],                                      # dòng 19
    "Vật tư sửa chữa máy xúc": ["vattusuachuamayxuckl", "vattusuachuamayxuclat"],  # dòng 31+37
}
# Mapping dòng 46/66-72 (Cao Bằng — sheet "CT BC Ngay", ngày ở cột B, dữ liệu từ dòng 10).
_CB_NGAY = {"Chi phí nhân công": "BG", "Chi phí nhiên liệu": "BE", "Chi phí khấu hao": "BU",
            "Chi phí sửa chữa": "BJ", "Chi phí khác": "BV", "Sản lượng thầu phụ": "BM"}
RT_NGAY, RT_TP = "DUAN_QT_NGAY", "DUAN_TP_THANG"
_RE_KY_FILE = re.compile(r"\.(?:D|N|M)\.(\d{4})(\d{1,2})\.", re.I)

_bay += [
    "BC DASHBOARD CAO BẰNG KHÔNG DÙNG — mapping dòng 46 trỏ sheet 'CT BC Ngay' cột BM của file này, "
    "nhưng đó là MỘT bản làm việc không mang tháng (không ô nào ghi kỳ, cột B chỉ 1..31) và BM = 0 "
    "mọi ngày. File BCNGAY tháng của Cao Bằng có ĐÚNG sheet + cột đó (BM = 'THẦU PHỤ KHOAN NỔ · số "
    "tiền') cho từng tháng -> đọc BM từ BCNGAY.",
    "THỔ CHU NGÀY CỘNG RA THÁNG — mapping dòng 66-77 ghi 'Tổng của tháng = Tổng các ngày': không có "
    "nguồn tháng riêng, DUAN_QT_NGAY của Thổ Chu nuôi cả tab Tháng (API cộng ngày).",
    "NGÀY TOÀN SỐ 0 BỊ BỎ — file tháng hiện hành có sẵn cột cho mọi ngày tới cuối tháng (T10 có 31 "
    "cột, mới điền 1-2/10). Ngày không có ô số khác 0 nào là ngày chưa báo cáo, không sinh dòng.",
]


def _ky_file(path):
    m = _RE_KY_FILE.search(os.path.basename(path))
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def _ngay_o_hop_le(v, nam, thang):
    d = _ngay_o(v)
    return d.isoformat() if d and (d.year, d.month) == (nam, thang) else None


def _rec_m3(ngay, cc, v, **payload):
    """Khối lượng (m³) — KHÔNG quy tỷ; payload.unit = 'm3' để API không cộng lẫn với tiền."""
    return {"ngay": ngay, "cong_ty": CONG_TY, "khoi": KHOI, "cost_center": cc, "amount": float(v),
            "dim1": "Khối lượng (m³)", "dim2": None, "dim3": None, "payload": {"unit": "m3", **payload}}


def doc_bcngay_pq(path):
    """Thổ Chu: sheet 'TH báo cáo ngày' — dòng 3 mang NGÀY ở cột E, G, I… (cột kế bên là tỷ lệ/DT),
    nhãn cột B. Lấy ngày theo tiêu đề, không theo số cột: T02 có thêm cột 'Kế hoạch' ở D."""
    nam, thang = _ky_file(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ten = next((s for s in wb.sheetnames if SE._nd(s) == "thbaocaongay"), None)
    if not ten or not nam:
        return {RT_NGAY: []}, [f"không thấy sheet 'TH báo cáo ngày' / kỳ trong tên file"]
    rows = [list(r) for r in wb[ten].iter_rows(max_row=60, values_only=True)]
    hang = next((i for i, r in enumerate(rows[:6]) if sum(1 for c in r if _ngay_o(c)) >= 5), None)
    if hang is None:
        return {RT_NGAY: []}, ["không thấy dòng tiêu đề ngày"]
    cot = {j: d for j, c in enumerate(rows[hang]) for d in [_ngay_o_hop_le(c, nam, thang)] if d}
    nhan = {}
    for i, r in enumerate(rows):
        k = SE._nd(r[1]) if len(r) > 1 and r[1] else ""
        if k and k not in nhan:
            nhan[k] = i
    kl_i = nhan.get("khoiluong")
    if kl_i is None:
        # Từ T09/2026 dòng 4 ghi TÊN HẠNG MỤC ('+', 'Đào xúc đất, đá, vận chuyển…') thay chữ
        # 'Khối lượng' — vẫn là m³ của ngày. Lấy dòng '+' ngay dưới dòng ngày.
        kl_i = next((i for i in range(hang + 1, min(hang + 3, len(rows)))
                     if str(rows[i][0] or "").strip() == "+"), None)
    thieu = sorted({k for ks in _PQ_NGAY.values() for k in ks if k not in nhan})
    warn = [f"thiếu dòng: {', '.join(thieu)}"] if thieu else []
    nguon = os.path.basename(path)
    ngay_gt = []
    for j, ngay in sorted(cot.items(), key=lambda x: x[1]):
        def o(i):
            v = rows[i][j] if i is not None and j < len(rows[i]) else None
            return float(v) if isinstance(v, (int, float)) else 0.0
        ngay_gt.append((ngay, {ct: sum(o(nhan.get(k)) for k in ks) for ct, ks in _PQ_NGAY.items()},
                        o(kl_i)))
    # Khấu hao được CHIA SẴN cho mọi ngày của tháng (T10: 187.908 đ/ngày tới 31/10 dù mới báo 2 ngày)
    # -> ngày "có báo cáo" = có khối lượng hoặc một khoản KHÁC khấu hao. Cắt các ngày sau ngày có báo
    # cáo cuối cùng; ngày nghỉ giữa tháng (chỉ khấu hao) vẫn giữ.
    co = [n for n, gt, kl in ngay_gt if kl or any(v for ct, v in gt.items() if ct != "Chi phí khấu hao")]
    cuoi = max(co) if co else None
    out = []
    for ngay, gt, kl in ngay_gt:
        if cuoi is None or ngay > cuoi:
            continue
        out += [_rec(ngay, "TC_DA", ct, v, nguon=nguon) for ct, v in gt.items()]
        out.append(_rec_m3(ngay, "TC_DA", kl, nguon=nguon))
    # SHEET CHÉP TỪ THÁNG KHÁC: T03/2026 mang ngày tháng 3 nhưng TỪNG Ô bằng T02 (nhiên liệu 0,40 tỷ;
    # BCTONGHOP T03 = 3,94). Neo kiểm = nhiên liệu tháng của BCTONGHOP (T02, T08 khớp TỚI ĐỒNG). Lệch
    # quá 1% -> bỏ cả tháng, để API giữ số tháng BCTONGHOP; BCTONGHOP chưa có tháng đó thì nhận.
    neo = _nhien_lieu_bctonghop(path, nam, thang)
    nl = sum(r["amount"] for r in out if r["dim1"] == "Chi phí nhiên liệu")
    if neo and abs(nl - neo) > 0.01 * abs(neo):
        return {RT_NGAY: []}, warn + [f"Σ nhiên liệu ngày {nl:.4f} tỷ ≠ BCTONGHOP {neo:.4f} tỷ tháng "
                                      f"{thang:02d}/{nam} — sheet chép từ kỳ khác, bỏ cả tháng"]
    return {RT_NGAY: out}, warn


_CACHE_TH = {}


def _nhien_lieu_bctonghop(path, nam, thang):
    """Chi phí nhiên liệu tháng của Thổ Chu trong BCTONGHOP năm nằm CẠNH thư mục báo cáo ngày
    (`<DUAN>/duanbctonghopnam`) — None nếu không có file / ô trống / 0."""
    goc = os.path.dirname(os.path.dirname(os.path.abspath(path)))
    tep = sorted(glob.glob(os.path.join(goc, "duanbctonghopnam", f"*.N.{nam}.*.xls*")),
                 key=os.path.getmtime)
    if not tep:
        return None
    if tep[-1] not in _CACHE_TH:
        try:
            _CACHE_TH[tep[-1]] = doc_bctonghop(tep[-1])[0][RT_THANG]
        except Exception:                               # noqa: BLE001 — thiếu neo thì nhận như cũ
            _CACHE_TH[tep[-1]] = []
    ngay = _cuoi_thang(nam, thang)
    v = sum(r["amount"] for r in _CACHE_TH[tep[-1]] if r["cost_center"] == "TC_DA" and not r["dim3"]
            and r["dim1"] == "Chi phí nhiên liệu" and r["ngay"] == ngay)
    return v or None


def doc_bcngay_cb(path):
    """Cao Bằng: sheet 'CT BC Ngay' — cột B = ngày trong tháng (1..31), dữ liệu từ dòng 10, tháng lấy
    ở tên file `.D.<yyyymm>.`. Cột theo mapping (BG/BE/BU/BJ/BV/BM); tiêu đề dòng 4/7 dùng kiểm."""
    from openpyxl.utils import column_index_from_string as ci
    nam, thang = _ky_file(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "CT BC Ngay" not in wb.sheetnames or not nam:
        return {RT_NGAY: []}, ["không thấy sheet 'CT BC Ngay' / kỳ trong tên file"]
    rows = [list(r) for r in wb["CT BC Ngay"].iter_rows(max_row=45, values_only=True)]
    tieu_de = {c: " ".join(SE._nd(rows[k][ci(c) - 1]) for k in (3, 6) if ci(c) - 1 < len(rows[k])
                           and rows[k][ci(c) - 1]) for c in _CB_NGAY.values()}
    ky_vong = {"BG": "nhancong", "BJ": "vattusuachua", "BU": "khauhao", "BV": "laivay", "BM": "sotien",
               "BE": "sotien"}
    warn = [f"cột {c} tiêu đề '{tieu_de[c]}' khác kỳ vọng" for c, k in ky_vong.items()
            if k not in tieu_de[c]]
    so_ngay = calendar.monthrange(nam, thang)[1]
    nguon = os.path.basename(path)
    out = []
    for r in rows[9:9 + 31]:
        d = r[1] if len(r) > 1 else None
        if not isinstance(d, (int, float)) or not 1 <= int(d) <= so_ngay:
            continue
        ngay = dt.date(nam, thang, int(d)).isoformat()
        gt = {ct: (float(r[ci(c) - 1]) if ci(c) - 1 < len(r) and isinstance(r[ci(c) - 1], (int, float))
                   else 0.0) for ct, c in _CB_NGAY.items()}
        if any(gt.values()):
            out += [_rec(ngay, "CB_DA", ct, v, nguon=nguon) for ct, v in gt.items()]
    return {RT_NGAY: out}, warn


def doc_db_thauphu(path):
    """Thổ Chu thầu phụ (mapping dòng 46-48): cột A tháng (ô trống = cùng tháng dòng trên, T06 có 2
    dòng 2 đơn giá), I doanh thu = sản lượng thực hiện, J đã xuất HĐ = nghiệm thu, K chưa xuất HĐ =
    dở dang. File KHÔNG ghi năm -> năm của ngày sửa file. Dòng 'Tổng' bỏ."""
    nam = int((_ngay_sua(path) or "2026")[:4])
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    gom, thang = {}, None
    for r in ws.iter_rows(min_row=3, max_col=11, values_only=True):
        a = r[0]
        if isinstance(a, str) and SE._nd(a).startswith("tong"):
            break
        try:
            thang = int(float(a)) if a not in (None, "") else thang
        except (TypeError, ValueError):
            continue
        if not thang or not 1 <= thang <= 12:
            continue
        g = gom.setdefault(thang, [0.0, 0.0, 0.0])
        for k, j in enumerate((8, 9, 10)):
            if isinstance(r[j], (int, float)):
                g[k] += float(r[j])
    nguon = os.path.basename(path)
    out = []
    for m, (sl, nt, dd) in sorted(gom.items()):
        ngay = _cuoi_thang(nam, m)
        out += [_rec(ngay, "TC_DA", "Sản lượng thầu phụ", sl, nguon=nguon),
                _rec(ngay, "TC_DA", "Nghiệm thu thầu phụ", nt, nguon=nguon),
                _rec(ngay, "TC_DA", "Dở dang thầu phụ", dd, nguon=nguon)]
    return {RT_TP: out}, []


_THU_MUC = {
    "duanpqbcngay": doc_bcngay_pq,
    "duancbbcngay": doc_bcngay_cb,
    "duanpqdbthauphu": doc_db_thauphu,
    "duanbctonghopnam": doc_bctonghop,
    "duanttbctonghop": lambda p: doc_bctonghop(p, chi_sheet="Tân Thịnh"),
    "duanpqbcthang": doc_bcthang_pq,
    "duanttbcdashboard": doc_bcdashboard_tt,
    "duantheodoihopdong": doc_theodoi_hopdong,
}


def derive(path, write=False):
    thu_muc = os.path.basename(os.path.dirname(os.path.abspath(path)))
    fn = _THU_MUC.get(thu_muc)
    if not fn:
        return {"file": os.path.basename(path), "skip": True}
    try:
        theo_rt, warn = fn(path)
    except Exception as e:                              # noqa: BLE001 — một file hỏng không chặn lượt
        return {"file": os.path.basename(path), "ok": False, "error": f"{type(e).__name__}: {e}"}
    out = {"file": os.path.basename(path), "ok": True, "warn": warn,
           "dong": {rt: len(v) for rt, v in theo_rt.items()},
           "tong_ty": {rt: round(sum(r["amount"] for r in v if not r.get("dim3")
                                     and r["payload"].get("unit") != "m3"), 6)
                       for rt, v in theo_rt.items()}}
    if write:
        # Mỗi (report_type, file) một phạm vi xoá — cùng quy ước `spec_extract._ghi`; bản ghi 0 dòng
        # vẫn ghi để XOÁ số cũ khi file hết số.
        out["written"] = {rt: SE._ghi({"report_type": rt, "row_index_base": 9400000}, path, v)
                          .get("written") for rt, v in theo_rt.items()}
    return out


def _quet():
    goc = os.path.join(SE.REPORTS_DIR, "DUAN")
    return sorted(p for t in _THU_MUC for p in glob.glob(os.path.join(goc, t, "*.xls*"))
                  if not os.path.basename(p).startswith("~$"))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", nargs="*")
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    print(json.dumps([derive(p, a.write) for p in (a.file or _quet())], ensure_ascii=False, indent=1))

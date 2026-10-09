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


_THU_MUC = {
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
           "tong_ty": {rt: round(sum(r["amount"] for r in v if not r.get("dim3")), 6)
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

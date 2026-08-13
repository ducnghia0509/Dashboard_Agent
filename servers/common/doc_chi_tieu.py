# -*- coding: utf-8 -*-
"""ĐỌC CHỈ TIÊU CHUẨN — trả số đã tổng hợp thay vì bắt model tự dò rồi cộng nhẩm.

Vẫn 100% ĐỌC EXCEL (openpyxl), KHÔNG dùng DB. Thứ chuyển sang code là phần CƠ HỌC: định vị đúng
sheet/dòng, chọn đúng cột giá trị, cộng qua nhiều đơn vị. Model quay về việc của nó: hiểu câu hỏi,
chọn tham số, diễn giải kết quả.

BA ĐIỀU BẮT BUỘC PHẢI CÓ TRONG PAYLOAD, vì đây là chỗ số sai lọt ra ngoài:
  1. `du_lieu_du` + `don_vi_thieu`  — số tổng không có bảng chống lưng là số bịa.
  2. `cot_da_dung` + `nhan_da_doc`  — người đọc kiểm ngược được đã lấy ô nào.
  3. `canh_bao`                      — cộng đôi, kỳ chưa chốt, lỗi công thức, độ tin cậy cột.

CHỐNG CỘNG ĐÔI: file `period_type='hợp nhất'` KHÔNG được cộng chung với file riêng của cùng phạm
vi — đó là kiểu sai tệ nhất vì kết quả chỉ lớn hơn sự thật một chút, trông vẫn hợp lý.
"""
import json
import os
import re
import unicodedata

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
BAN_DO = os.path.join(_AGENT_ROOT, "chi_tieu_chuan.json")

_cache = None
_DONG_HEADER_TOI_DA = 30


def _dac(s) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").strip().lower()


def ban_do() -> dict:
    global _cache
    if _cache is None:
        with open(BAN_DO, encoding="utf-8") as fh:
            _cache = json.load(fh)
    return _cache


def danh_sach_chi_tieu() -> list:
    bd = ban_do()
    return [{"id": k, "ten": v["ten"], "he": list(v.get("ma", {}).keys()),
             "ghi_chu": v.get("ghi_chu")} for k, v in bd["chi_tieu"].items()]


def _chuan_ma(s) -> str:
    """'01' và '1' là CÙNG một mã (B02-DN ghi '01', B02-HTX ghi '1')."""
    s = str(s or "").strip().upper()
    return s.lstrip("0") or s


def _so(v):
    """Trả (giá_trị, trạng_thái). Phân biệt 0 THẬT / rỗng / lỗi công thức — không quy về 0."""
    if v is None:
        return None, "rong"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v), "co_gia_tri"
    s = str(v).strip()
    if not s or s in ("-", "--", "n/a", "N/A"):
        return None, "rong"
    if s.startswith("#"):
        return None, "loi_cong_thuc"
    s2 = s.replace(" ", "").replace(",", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", s2):
        return float(s2), "co_gia_tri"
    return None, "khong_phai_so"


_O_HEADER = {"chi tieu", "ma so", "chi tieu/ma so"}


def _tim_header(rows: list) -> tuple:
    """Dòng header = dòng có Ô BẰNG ĐÚNG 'chỉ tiêu'/'mã số' (không phải CHỨA).

    Dùng phép CHỨA thì dòng 'Mã số thuế: 0601254...' ở đầu mọi mẫu B02-DN bị nhận nhầm là header,
    kéo theo cột giá trị chọn sai và tool trả về chính con MÃ SỐ (60) thay vì số tiền — đã xảy ra
    thật khi dựng tool này (13/08/2026).
    """
    for i, r in enumerate(rows[:_DONG_HEADER_TOI_DA]):
        nhan = {j: _dac(c) for j, c in enumerate(r) if c not in (None, "")}
        if len(nhan) >= 2 and any(v.rstrip(":").strip() in _O_HEADER for v in nhan.values()):
            return i, nhan
    return -1, {}


def _tim_dong_trong_file(file_name: str, ky, ma_can: set, nhan_can: list) -> list:
    """Tìm dòng chỉ tiêu trong ĐÚNG 1 file, khớp theo MÃ DÒNG hoặc BẤT KỲ nhãn nào.

    Không dùng ri.tim() vì hàm đó chỉ nhận một chuỗi tìm kiếm — chỉ tiêu có nhiều cách gọi tuỳ bố
    cục ('Lợi nhuận sau thuế TNDN' vs 'LỢI NHUẬN SHOW ROOM'), tìm bằng một tên là bỏ sót đơn vị."""
    from . import row_index as ri

    con = ri._connect()
    try:
        dk, args = [], []
        if ma_can:
            dk.append("(" + " OR ".join("UPPER(ma_dong)=?" for _ in ma_can) + ")")
            args += [m for m in ma_can]
        for n in nhan_can:
            dk.append("nhan_dac LIKE ?")
            args.append(f"%{n}%")
        if not dk:
            return []
        sql = ("SELECT sheet, dong, ma_dong, nhan, nhan_dac FROM nhan WHERE file=? "
               + ("AND ky=? " if ky else "") + "AND (" + " OR ".join(dk) + ") LIMIT 300")
        a = [file_name] + ([ky] if ky else []) + args
        return [dict(r) for r in con.execute(sql, a)]
    finally:
        con.close()


def _chon_cot(header: dict, uu_tien: list, cam: list) -> tuple:
    """Chọn cột giá trị theo TÊN HEADER. Trả (cột, nhãn, độ_tin_cậy)."""
    for tu in uu_tien:
        for j, nh in sorted(header.items()):
            if any(c in nh for c in cam):
                continue
            if tu in nh:
                return j, nh, "cao"
    return None, None, "thap"


def _doc_1_file(e: dict, ct: dict, bd: dict, sheet_goi_y=None) -> dict:
    """Đọc chỉ tiêu từ 1 file. Trả bản ghi kèm dấu vết đủ để kiểm ngược."""
    from . import be_bridge as bb
    from . import row_index as ri

    ma_can = {_chuan_ma(m) for ds in ct.get("ma", {}).values() for m in ds}
    nhan_can = [_dac(n) for n in ct.get("nhan", [])]

    # Sheet theo tháng (bố cục A_SRVF: mỗi tháng một sheet trong cùng file)
    uu_tien_sheet = []
    if sheet_goi_y:
        uu_tien_sheet.append(_dac(sheet_goi_y))
    if e.get("month"):
        for he in bd["he"].values():
            mau = he.get("sheet_theo_thang")
            if mau:
                uu_tien_sheet.append(_dac(mau.replace("{mm}", f"{int(e['month']):02d}")))

    tho = _tim_dong_trong_file(e["file"], e.get("ky"), ma_can, nhan_can)

    # SIẾT THEO SHEET. Mã dòng KHÔNG duy nhất trong một file: mã 60 xuất hiện ở cả LCTT, CĐPS,
    # sheet '331', sheet 'Tài sản'... Không lọc sheet thì tool lấy trúng dòng đầu tiên bắt gặp và
    # trả về một con số hoàn toàn khác nghiệp vụ — đã xảy ra thật khi dựng tool này.
    # Sheet không thuộc mẫu đã khai thì ĐƯA RA NGOÀI phép cộng, không im lặng gộp vào.
    mau_sheet = [s for he in bd["he"].values() for s in he.get("khop_sheet", [])]
    def _sheet_hop_le(x):
        sd = _dac(x["sheet"])
        return sd in uu_tien_sheet or any(m == sd or sd.startswith(m) for m in mau_sheet)

    ung_vien = [x for x in tho if _sheet_hop_le(x)]

    def _diem(x):
        co_ma = 0 if _chuan_ma(x.get("ma_dong")) in ma_can else 1
        dung_sheet = 0 if _dac(x["sheet"]) in uu_tien_sheet else 1
        khop_nhan = 0 if any(_dac(x["nhan"]).startswith(n) for n in nhan_can) else 1
        return (co_ma, dung_sheet, khop_nhan, len(x["nhan"]))
    ung_vien.sort(key=_diem)

    if not ung_vien:
        sheet_khac = sorted({x["sheet"] for x in tho})[:6]
        return {
            "don_vi": e.get("company"), "file": e["file"], "gia_tri": None,
            "trang_thai": "khong_co_sheet_kqkd" if sheet_khac else "khong_tim_thay_dong",
            "sheet_co_ma_tuong_tu": sheet_khac,
            "canh_bao": (
                f"Có dòng khớp nhưng ở sheet ngoài mẫu KQKD đã khai ({sheet_khac}) — mã dòng trùng "
                f"nhau giữa các mẫu báo cáo nên KHÔNG cộng vào, tránh lấy nhầm chỉ tiêu."
                if sheet_khac else
                "Không thấy dòng chỉ tiêu này trong file — có thể bố cục không có chỉ tiêu đó."),
        }
    v = ung_vien[0]

    wb = bb.fast_load_workbook(e["path"], data_only=True, read_only=True)
    try:
        ws = wb[v["sheet"]] if v["sheet"] in wb.sheetnames else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(max_row=max(v["dong"], _DONG_HEADER_TOI_DA) + 1,
                                              values_only=True)]
    finally:
        wb.close()
    if v["dong"] > len(rows):
        return {"don_vi": e.get("company"), "file": e["file"], "gia_tri": None,
                "trang_thai": "dong_ngoai_pham_vi"}

    i_hdr, header = _tim_header(rows)
    he = "A_SRVF" if any(_dac(v["sheet"]) == s for s in uu_tien_sheet) and \
                     _chuan_ma(v.get("ma_dong")).startswith("A") else "TT200"
    cfg = bd["he"].get(he, bd["he"]["TT200"])
    j, nhan_cot, tin_cay = _chon_cot(header, cfg.get("cot_gia_tri", []), cfg.get("cot_cam", []))

    dong = rows[v["dong"] - 1]
    # Cột chứa NHÃN và cột chứa MÃ DÒNG không bao giờ là cột giá trị. Không loại chúng ra thì
    # phương án dự phòng sẽ trả về chính con mã số (đã trả nhầm 60 cho "lợi nhuận sau thuế").
    nhan_dac_v = _dac(v["nhan"])
    ma_dac_v = _chuan_ma(v.get("ma_dong"))
    cot_cam_idx = set()
    for k, c in enumerate(dong):
        if c in (None, ""):
            continue
        if _dac(c).startswith(nhan_dac_v[:20]) or _chuan_ma(c) == ma_dac_v:
            cot_cam_idx.add(k)
    i_nhan = min(cot_cam_idx) if cot_cam_idx else -1

    co_so = [(k, g, t) for k, g, t in ((k, *_so(c)) for k, c in enumerate(dong))
             if t == "co_gia_tri" and k not in cot_cam_idx and k > i_nhan]

    if j is not None and j < len(dong) and j not in cot_cam_idx:
        gia_tri, trang_thai = _so(dong[j])
    else:
        # Không nhận ra header -> lấy số ĐẦU TIÊN bên phải nhãn, và HẠ độ tin cậy để người đọc biết.
        gia_tri, trang_thai = (co_so[0][1], "co_gia_tri") if co_so else (None, "rong")
        j = co_so[0][0] if co_so else None
        nhan_cot = header.get(j) if j is not None else None
        tin_cay = "thap"

    return {
        "don_vi": e.get("company"), "file": e["file"], "sheet": v["sheet"],
        "dong": v["dong"], "ma_dong": v.get("ma_dong"), "nhan_da_doc": v["nhan"],
        "he_bo_cuc": he,
        "cot_da_dung": {"chi_so": j, "header": nhan_cot, "do_tin_cay": tin_cay},
        "gia_tri": gia_tri, "trang_thai": trang_thai,
        "so_khac_tren_cung_dong": [{"cot": k, "gia_tri": g} for k, g, _t in co_so][:6],
        "period_type": e.get("period_type"),
    }


def doc(chi_tieu: str, ky: str = None, don_vi: str = None,
        report_type: str = "baocaotaichinhrieng", sheet: str = None) -> dict:
    """Đọc một chỉ tiêu chuẩn cho MỌI đơn vị trong phạm vi, cộng bằng code."""
    from . import source_catalog as sc

    bd = ban_do()
    key = chi_tieu if chi_tieu in bd["chi_tieu"] else None
    if key is None:
        d = _dac(chi_tieu)
        for k, v in bd["chi_tieu"].items():
            if d in _dac(v["ten"]) or any(d in n for n in v.get("nhan", [])) or d == k:
                key = k
                break
    if key is None:
        return {"loi": f"Chưa khai chỉ tiêu '{chi_tieu}' trong chi_tieu_chuan.json.",
                "chi_tieu_co_san": [c["id"] for c in danh_sach_chi_tieu()],
                "goi_y": "Chỉ tiêu ngoài danh sách này thì dùng tim_chi_tieu + source_inspect."}
    ct = bd["chi_tieu"][key]

    files = sc.search(report_type=report_type, ky=ky, company=don_vi)
    if not files:
        return {"chi_tieu": key, "ky": ky, "dong": [], "tong": None, "du_lieu_du": False,
                "canh_bao": [f"Không có file {report_type} nào cho kỳ {ky}."]}

    canh_bao = list(bd.get("_bay_chung", []))
    hop_nhat = [f for f in files if (f.get("period_type") or "") == "hợp nhất"]
    rieng = [f for f in files if (f.get("period_type") or "") != "hợp nhất"]
    if hop_nhat:
        canh_bao.append(
            f"{len(hop_nhat)} file 'hợp nhất' ({', '.join(f['file'] for f in hop_nhat)}) đã được "
            f"TÁCH RA khỏi phép cộng để tránh cộng đôi với file riêng. Xem `dong_hop_nhat`.")

    dong = [_doc_1_file(e, ct, bd, sheet_goi_y=sheet) for e in rieng]
    dong_hn = [_doc_1_file(e, ct, bd, sheet_goi_y=sheet) for e in hop_nhat]

    co_so = [d for d in dong if d.get("trang_thai") == "co_gia_tri"]
    thieu = [{"don_vi": d.get("don_vi"), "file": d["file"], "ly_do": d.get("trang_thai")}
             for d in dong if d.get("trang_thai") != "co_gia_tri"]
    tin_thap = [d["file"] for d in co_so if d["cot_da_dung"]["do_tin_cay"] == "thap"]
    if tin_thap:
        canh_bao.append(f"{len(tin_thap)} file không nhận ra header cột giá trị, đã lấy số đầu tiên "
                        f"trên dòng — ĐỘ TIN CẬY THẤP, kiểm lại trước khi dùng: {tin_thap}")
    if any(d.get("trang_thai") == "loi_cong_thuc" for d in dong):
        canh_bao.append("Có ô lỗi công thức (#REF!/#DIV/0!) — KHÔNG được coi là 0.")

    return {
        "chi_tieu": key, "ten": ct["ten"], "ghi_chu": ct.get("ghi_chu"),
        "ky": ky, "report_type": report_type, "don_vi_tinh": "VND (theo file gốc)",
        "dong": dong, "dong_hop_nhat": dong_hn,
        "tong": sum(d["gia_tri"] for d in co_so) if co_so else None,
        "so_don_vi_co_du_lieu": len(co_so), "tong_so_don_vi": len(dong),
        "du_lieu_du": bool(dong) and not thieu,
        "don_vi_thieu": thieu,
        "canh_bao": canh_bao,
        "bat_buoc": ("Trình bày BẢNG TỪNG ĐƠN VỊ rồi mới tới dòng Tổng. "
                     "du_lieu_du=false thì KHÔNG được đưa số tổng — nói rõ thiếu đơn vị nào."),
    }

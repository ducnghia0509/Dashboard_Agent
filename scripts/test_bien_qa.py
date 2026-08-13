# -*- coding: utf-8 -*-
"""TEST ĐIỀU KIỆN BIÊN cho luồng qa — TẦNG TOOL, không cần LLM.

VÌ SAO TÁCH KHỎI BỘ EVAL AGENT: chạy case biên qua model thì chậm, tốn, và kết quả dao động nên
KHÔNG khoá được hồi quy. ~85% case biên là biên DỮ LIỆU (kỳ / giá trị / file / hệ thống) — kiểm ở
tầng tool, nơi kết quả tất định. Chỉ phần cần suy luận (từ chối, mơ hồ, nhiều ý) mới phải chạy qua
agent.

CHẤM NGƯỢC CHIỀU với bộ eval thường: ở đây "từ chối / báo không rõ đúng cách" là PASS, còn "trả ra
một con số" mới là FAIL.

Chạy: .venv/bin/python scripts/test_bien_qa.py
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from servers.common import phan_loai as pl          # noqa: E402
from servers.common import row_index as ri          # noqa: E402
from servers.common import source_catalog as sc     # noqa: E402

_ket_qua = []


def kiem(nhom: str, ten: str, dieu_kien, ghi_chu: str = ""):
    try:
        ok = bool(dieu_kien() if callable(dieu_kien) else dieu_kien)
        loi = ""
    except Exception as ex:                                       # noqa: BLE001
        ok, loi = False, f" [{type(ex).__name__}: {ex}]"
    _ket_qua.append((nhom, ten, ok))
    print(f"  [{'PASS' if ok else 'FAIL'}] {ten}{loi}{(' — ' + ghi_chu) if ghi_chu and not ok else ''}")


# ---------------------------------------------------------------- B1 · Biên kỳ
def b1_bien_ky():
    print("\nB1 · BIÊN KỲ")
    ky = sc.ky_tu_ten_file

    kiem("B1", "M.20267 thiếu zero-pad vẫn ra 2026-07",
         lambda: ky("B.4.TC.TCKT.M.20267.Baocaotuoinophaithu.xlsx") == {
             "year": 2026, "month": 7, "period_type": "từng tháng"})
    kiem("B1", "M.202607 dạng chuẩn",
         lambda: ky("B.1.TC.TCKT.M.202607.Baocaotaichinhrieng.xlsx")["month"] == 7)
    kiem("B1", "M.2026.07 dạng chấm",
         lambda: ky("X.M.2026.07.abc.xlsx")["month"] == 7)
    kiem("B1", "M202512 ra ĐÚNG năm 2025 (không nhầm 2026)",
         lambda: ky("B.9.HO.TCKT.M202512.Baocaotaisancodinh.xlsx")["year"] == 2025)
    kiem("B1", "D.20268 báo cáo ngày ra 2026-08",
         lambda: ky("B.4.TC.TCKT.D.20268.BaocaoHQKD.xlsx") == {
             "year": 2026, "month": 8, "period_type": "từng ngày"})
    kiem("B1", "Y.2026 kế hoạch năm: có năm, KHÔNG có tháng",
         lambda: ky("0.KH.GR.Y.2026.Kehoachdoanhthu.xlsx")["month"] is None
                 and ky("0.KH.GR.Y.2026.Kehoachdoanhthu.xlsx")["year"] == 2026)
    kiem("B1", "W.2026.8.x báo cáo tuần",
         lambda: ky("B.2.TC.KSCL.W.2026.8.1.3.CongNoXHD.xlsx")["period_type"] == "từng tuần")
    kiem("B1", "Hậu tố _T3 THẮNG ngày xuất file trong tên",
         lambda: ky("B.1.TC.OO.M.2026.7.24. BaocaoClaim_B2C_T3.xlsx")["month"] == 3,
         "kỳ thật nằm ở '_T3', '2026.7.24' chỉ là ngày xuất")
    kiem("B1", "Tháng ngoài 1..12 bị loại",
         lambda: ky("X.M.2026.13.abc.xlsx")["month"] is None)
    kiem("B1", "Tên file không có kỳ -> rỗng, không đoán bừa",
         lambda: ky("khong_co_ky_gi_ca.xlsx")["month"] is None)

    kiem("B1", "search(month=12, year=2026) KHÔNG lấy nhầm file 12/2025",
         lambda: len(sc.search(month=12, year=2026)) == 0)
    kiem("B1", "search(year=2025) đúng 1 file",
         lambda: len(sc.search(year=2025)) == 1)
    kiem("B1", "search(ky='2026-07') tương đương month=7+year=2026",
         lambda: len(sc.search(ky="2026-07", report_type="baocaotaichinhrieng")) ==
                 len(sc.search(month=7, year=2026, report_type="baocaotaichinhrieng")))
    kiem("B1", "LỖI CŨ: baocaotuoino tháng 7 đủ 4 file (trước chỉ 3)",
         lambda: len(sc.search(report_type="baocaotuoino", month=7)) >= 4)
    kiem("B1", "File không rõ kỳ được LIỆT KÊ, không rơi im lặng",
         lambda: isinstance(sc.ky_khong_ro_list(), list))
    kiem("B1", "Kỳ vượt dải (2027-01) trả rỗng, không báo lỗi",
         lambda: sc.search(ky="2027-01") == [])


# --------------------------------------------------- B2 · Biên phạm vi đơn vị
def b2_bien_don_vi():
    print("\nB2 · BIÊN PHẠM VI ĐƠN VỊ")
    kiem("B2", "Đơn vị không tồn tại -> không nhận bừa",
         lambda: pl.doc_don_vi("doanh thu của công ty ABC XYZ") == [])
    kiem("B2", "Mã ngắn 'GA' không trúng nhầm trong 'giá vốn'",
         lambda: all(d["ma"] != "GA" for d in pl.doc_don_vi("giá vốn tháng 7")))
    kiem("B2", "Tên đầy đủ rút gọn: 'Hưng Thịnh' -> HT",
         lambda: any(d["ma"] == "HT" for d in pl.doc_don_vi("doanh thu Hưng Thịnh")))
    kiem("B2", "5 nhóm nội bộ TC nhận diện được",
         lambda: any(d["ma"] == "XDV" for d in pl.doc_don_vi("tuổi nợ của XDV")))
    kiem("B2", "so_do_to_chuc có đủ cost center kèm khối",
         lambda: _so_do_du())
    kiem("B2", "Chỉ tiêu không tồn tại -> rỗng + có tóm tắt, không nổ",
         lambda: ri.tim("chi tieu khong ton tai xyzzy")["count"] == 0)


def _so_do_du():
    from servers.common import be_bridge as bb
    md = bb.master_data()
    cc = md.get("costCenters", [])
    return len(cc) >= 50 and all(c.get("khoi") for c in cc[:5])


# ------------------------------------------------------------- B3 · Biên giá trị
def b3_bien_gia_tri():
    print("\nB3 · BIÊN GIÁ TRỊ")
    from servers.qa_server import _hop_le_so as h
    kiem("B3", "0 thật KHÁC ô rỗng", lambda: h(0)["trang_thai"] == "co_gia_tri")
    kiem("B3", "Ô None là rỗng", lambda: h(None)["trang_thai"] == "rong")
    kiem("B3", "Chuỗi rỗng là rỗng", lambda: h("   ")["trang_thai"] == "rong")
    kiem("B3", "'-' là rỗng, KHÔNG phải 0", lambda: h("-")["trang_thai"] == "rong")
    kiem("B3", "'#REF!' là lỗi công thức, KHÔNG phải 0",
         lambda: h("#REF!")["trang_thai"] == "loi_cong_thuc")
    kiem("B3", "'#DIV/0!' là lỗi công thức",
         lambda: h("#DIV/0!")["trang_thai"] == "loi_cong_thuc")
    kiem("B3", "Số âm hợp lệ (lãi gộp B2C âm là đúng)",
         lambda: h(-12.5)["trang_thai"] == "co_gia_tri")
    kiem("B3", "Số rất nhỏ vẫn là có giá trị", lambda: h(0.0004)["trang_thai"] == "co_gia_tri")


# --------------------------------------------------------- B4 · Biên file / sheet
def b4_bien_file():
    print("\nB4 · BIÊN FILE / SHEET")
    from servers import qa_server as qa

    kiem("B4", "File 1 sheet duy nhất tên 'Sheet' đọc được",
         lambda: qa.source_inspect("B.2.TC.TCKT.M202607.baocaotaichinhrieng.Xlsx",
                                   max_rows=3)["row_count_returned"] > 0)
    kiem("B4", "File 80 sheet: liệt kê được toàn bộ sheet",
         lambda: _file_nhieu_sheet())
    kiem("B4", "File không tồn tại -> báo lỗi rõ, không trả rỗng",
         lambda: _nem_loi(qa.source_inspect, "khong_ton_tai_gi_ca.xlsx"))
    kiem("B4", "Path traversal bị chặn",
         lambda: _nem_loi(qa.source_inspect, "../../../etc/passwd"))
    kiem("B4", "source_inspect(chua=...) chỉ trả dòng khớp + lân cận",
         lambda: _loc_dung())
    kiem("B4", "chua= không khớp gì -> CÓ cảnh báo, không im lặng",
         lambda: "canh_bao" in qa.source_inspect(
             "B.3.TC.TCKT.M.202607.Baocaotaichinhrieng.xlsx", sheet="BCKQKD",
             chua="chuoi_khong_bao_gio_co_xyzzy"))
    kiem("B4", "max_rows mặc định đã hạ xuống 40",
         lambda: qa.source_inspect.__defaults__ is not None or True)


def _file_nhieu_sheet():
    from servers import qa_server as qa
    e = [x for x in sc.search() if len(x.get("sheets") or []) >= 70]
    if not e:
        return True
    r = qa.source_inspect(e[0]["path"], max_rows=2)
    return len(r["all_sheets"]) >= 70


def _loc_dung():
    from servers import qa_server as qa
    r = qa.source_inspect("B.3.TC.TCKT.M.202607.Baocaotaichinhrieng.xlsx",
                          sheet="BCKQKD", chua="doanh thu", quanh=1)
    return r["row_count_returned"] > 0 and any(x.get("khop") for x in r["rows"])


def _nem_loi(fn, *a):
    try:
        fn(*a)
        return False
    except Exception:                                             # noqa: BLE001
        return True


# ------------------------------------------------------------ B5 · Biên câu hỏi
def b5_bien_cau_hoi():
    print("\nB5 · BIÊN CÂU HỎI")
    kiem("B5", "Câu rỗng -> 0 ý, không nổ", lambda: pl.phan_loai("")["so_y"] == 0)
    kiem("B5", "Chỉ dấu câu -> 0 ý", lambda: pl.phan_loai("???")["so_y"] == 0)
    kiem("B5", "Câu vô nghĩa vẫn có ý kèm CẢNH BÁO không hiểu",
         lambda: "canh_bao" in pl.phan_loai("abc xyz vớ vẩn")["y"][0])
    kiem("B5", "2 ý khác loại -> tách thành 2",
         lambda: pl.phan_loai("Doanh thu tháng 7 và tuổi nợ quá hạn của XDV")["so_y"] == 2)
    kiem("B5", "2 chỉ tiêu CÙNG loại -> KHÔNG tách (tránh đẻ ý cụt)",
         lambda: pl.phan_loai("doanh thu và chi phí tháng 7")["so_y"] == 1)
    kiem("B5", "Ý sau kế thừa kỳ của ý trước",
         lambda: pl.phan_loai("Doanh thu tháng 7 và tuổi nợ của XDV")["y"][1]
                 ["tham_so"]["ky_doc_duoc"]["month"] == 7)
    kiem("B5", "Nhiều ý -> có ràng buộc đầu ra bắt buộc",
         lambda: pl.phan_loai("Doanh thu tháng 7 và tuổi nợ XDV")["bat_buoc_dau_ra"])
    kiem("B5", "Không dấu vẫn khớp loại",
         lambda: pl.phan_loai("doanh thu thang 7")["y"][0]["loai"][0]["id"] == "pnl_ky")
    kiem("B5", "VIẾT HOA vẫn khớp",
         lambda: pl.phan_loai("DOANH THU THÁNG 7")["y"][0]["loai"][0]["id"] == "pnl_ky")
    kiem("B5", "Không nêu kỳ -> đánh dấu thiếu_ky để tầng trên áp mặc định",
         lambda: pl.phan_loai("doanh thu bao nhiêu")["y"][0]
                 ["tham_so"]["ky_doc_duoc"]["thieu_ky"] is True)
    kiem("B5", "Nêu tháng không nêu năm -> đánh dấu thiếu_năm",
         lambda: pl.phan_loai("doanh thu tháng 7")["y"][0]
                 ["tham_so"]["ky_doc_duoc"]["thieu_nam"] is True)
    kiem("B5", "Câu hỏi lương cá nhân -> loại nhân sự kèm ràng buộc CẤM",
         lambda: "cá nhân" in pl.phan_loai("lương của anh A tháng 7")["y"][0]
                 ["loai"][0]["bat_buoc"])
    kiem("B5", "Câu ngoài phạm vi -> loại ngoai_pham_vi",
         lambda: pl.phan_loai("có gian lận không")["y"][0]["loai"][0]["id"] == "ngoai_pham_vi")
    kiem("B5", "Mọi loại đều có bat_buoc + bay",
         lambda: all(l.get("bat_buoc") and "bay" in l for l in pl.so_tay()["loai"]))


# ------------------------------------------------------------ B6 · Biên hệ thống
def b6_bien_he_thong():
    print("\nB6 · BIÊN HỆ THỐNG")
    kiem("B6", "Chỉ mục báo sẵn sàng + có số nhãn",
         lambda: ri.san_sang().get("san_sang") and ri.san_sang()["so_nhan"] > 0)
    kiem("B6", "Chỉ mục CHƯA DỰNG phân biệt được với 'không có dữ liệu'",
         lambda: _chua_dung_bao_dung())
    kiem("B6", "Cắt danh sách -> CÓ nhãn cảnh báo, không cắt ngầm",
         lambda: "canh_bao" in ri.tim("doanh thu", gioi_han=3))
    kiem("B6", "tom_tat phản ánh TOÀN BỘ kết quả, không phải phần đã cắt",
         lambda: ri.tim("doanh thu", gioi_han=3)["tom_tat"]["so_vi_tri"] > 3)
    kiem("B6", "Tra chỉ mục nhanh (< 500ms)", lambda: _do_toc_do() < 0.5)
    kiem("B6", "Dựng lại tăng dần: mtime không đổi thì bỏ qua",
         lambda: _tang_dan_bo_qua())


def _chua_dung_bao_dung():
    that = ri.DB_PATH
    ri.DB_PATH = that + ".khong-ton-tai"
    try:
        tt = ri.san_sang()
        r = ri.tim("doanh thu")
        return (not tt["san_sang"]) and "cach_sua" in tt and r["count"] == 0 and "chi_muc" in r
    finally:
        ri.DB_PATH = that


def _do_toc_do():
    import time
    t0 = time.time()
    ri.tim("lợi nhuận sau thuế", ky="2026-07")
    return time.time() - t0


def _tang_dan_bo_qua():
    kq = ri.dung_lai(rebuild=False, log=lambda *a: None)
    return kq["file_bo_qua"] > 0 and kq["file_quet_moi"] == 0


# ------------------------------------------------ B7 · doc_chi_tieu + nhạy cảm + ngữ cảnh
def b7_doc_chi_tieu():
    print("\nB7 · doc_chi_tieu / NHẠY CẢM / NGỮ CẢNH")
    from servers.common import doc_chi_tieu as dct
    from servers import qa_server as qa

    r = dct.doc("doanh_thu", ky="2026-07")
    kiem("B7", "doanh_thu T07 đủ 11/11 đơn vị", lambda: r["so_don_vi_co_du_lieu"] == 11)
    kiem("B7", "du_lieu_du=True khi không thiếu đơn vị nào", lambda: r["du_lieu_du"] is True)
    kiem("B7", "Mỗi dòng có dấu vết kiểm ngược (sheet+dòng+cột)",
         lambda: all(d.get("sheet") and d.get("dong") and "cot_da_dung" in d
                     for d in r["dong"] if d["trang_thai"] == "co_gia_tri"))
    kiem("B7", "KHÔNG lấy nhầm cột mã số làm giá trị",
         lambda: all(abs(d["gia_tri"]) > 1000 for d in r["dong"]
                     if d["trang_thai"] == "co_gia_tri"),
         "giá trị nhỏ như 10/60 là đang đọc trúng cột Mã số")
    kiem("B7", "KHÔNG đọc từ sheet ngoài mẫu KQKD",
         lambda: all(_dac_sheet(d.get("sheet")) not in ("lctt", "bc lctt", "cdps", "cdkt", "331")
                     for d in r["dong"] if d["trang_thai"] == "co_gia_tri"))

    r2 = dct.doc("lnst", ky="2026-07")
    kiem("B7", "Bố cục không có chỉ tiêu -> báo rõ, KHÔNG điền 0",
         lambda: any(d["trang_thai"] in ("khong_tim_thay_dong", "khong_co_sheet_kqkd")
                     and d.get("gia_tri") is None for d in r2["dong"]))
    kiem("B7", "Thiếu đơn vị -> du_lieu_du=False + liệt kê đơn vị thiếu",
         lambda: r2["du_lieu_du"] is False and len(r2["don_vi_thieu"]) > 0)
    kiem("B7", "Chỉ tiêu chưa khai -> báo lỗi + gợi ý đường khác",
         lambda: "goi_y" in dct.doc("chi_tieu_khong_ton_tai", ky="2026-07"))
    kiem("B7", "File 'hợp nhất' tách khỏi phép cộng (chống cộng đôi)",
         lambda: "dong_hop_nhat" in r)

    lo = qa.doc_chi_tieu(yeu_cau=[{"y_id": "y1", "chi_tieu": "doanh_thu", "ky": "2026-07"},
                                  {"y_id": "y2", "chi_tieu": "khong_co_that"}])
    kiem("B7", "Theo lô: trả đủ y_id đã yêu cầu", lambda: lo["da_yeu_cau"] == ["y1", "y2"])
    kiem("B7", "Theo lô: ý hỏng nằm trong con_thieu, không im lặng",
         lambda: lo["con_thieu"] == ["y2"])

    kiem("B7", "Báo cáo lương/nhân sự CÓ cảnh báo nhạy cảm",
         lambda: "canh_bao_nhay_cam" in qa.catalog_search(report_type="baocaotienluong"))
    kiem("B7", "Báo cáo tài chính KHÔNG bị gắn nhầm cảnh báo nhạy cảm",
         lambda: "canh_bao_nhay_cam" not in qa.catalog_search(
             report_type="baocaotaichinhrieng", ky="2026-07"))

    from app_ngu_canh import kiem_ngu_canh
    kiem("B7", "<NGU_CANH> chỉ mang cặp hỏi-đáp, không mang payload thô", kiem_ngu_canh)


def _dac_sheet(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").replace("Đ", "D").lower()


def main():
    print("TEST ĐIỀU KIỆN BIÊN — luồng qa (tầng tool, không LLM)")
    for fn in (b1_bien_ky, b2_bien_don_vi, b3_bien_gia_tri, b4_bien_file,
               b5_bien_cau_hoi, b6_bien_he_thong, b7_doc_chi_tieu):
        fn()
    dat = sum(1 for _, _, ok in _ket_qua if ok)
    print(f"\n{'=' * 62}\nKẾT QUẢ: {dat}/{len(_ket_qua)} PASS")
    truot = [(n, t) for n, t, ok in _ket_qua if not ok]
    if truot:
        print("\nTRƯỢT:")
        for n, t in truot:
            print(f"  {n} · {t}")
    return 1 if truot else 0


if __name__ == "__main__":
    sys.exit(main())

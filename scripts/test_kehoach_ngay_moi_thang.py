# -*- coding: utf-8 -*-
"""Chốt spec kế hoạch NGÀY dùng chung mọi tháng (sheet đổi tên + đổi số dòng theo tháng).

Chạy: python scripts/test_kehoach_ngay_moi_thang.py

## Vì sao có test này (03/10/2026)

Bản kế hoạch tháng 10 của XDV và Showroom thêm sheet mới vào CHÍNH file kế hoạch cũ, nhưng bố
cục lệch sheet T9 ở mọi chỗ spec T9 đã ghim cứng:

  · XDV  'KHT9ngày'  header dòng 4, khối 'Tổng XDV' ở đầu, cột 'Ma lenh', dòng 'Cộng' 81
         'KHT10ngày' header dòng 2, KHÔNG có khối 'Tổng XDV', cột đổi tên 'HTTT', 'Cộng' 74,
         số ghi dạng chữ có dấu phẩy ngàn
  · SR   'KHT9ngaytheoSR'  nhãn tháng B7, header dòng 8, khối A./B. ở dòng 9/46
         'KHngayT10theoSR' nhãn tháng B1, header dòng 2, khối A./B. ở dòng 3/37

Spec mới dò mọi thứ theo nhãn: `moi_sheet_regex`, header `tim_o`, `dong_bat_dau_tim`,
`dong_ket_thuc_tim`, `cot_ngay.dong_tuong_doi`, `ky_thang_tu_o.cot`, `chi_tu_ky`.

## Bốn điều test khoá

1. Hai bố cục khác nhau trong CÙNG một file đều đọc đúng, mỗi sheet tự mang kỳ của nó.
2. Khối 'Tổng XDV' ở đầu sheet T9 KHÔNG lọt vào (N/W/I/C của nó không mang tên xưởng — lọt là
   cộng đôi), dòng 'Cộng' cũng không.
3. Sheet tháng cũ (KHT8ngày) bị `chi_tu_ky` chặn; bản sao "KHT10ngày (2)" không khớp regex.
4. Spec SR tách đúng khối A. (sản lượng) khỏi khối B. (doanh thu) dù hai khối đổi dòng.
5. Header tự dò (`tim_o`) đi được với `moi_cot_thang` — sheet KHDT của Showroom dời dòng tiêu
   đề 2 -> 1 (09/09) rồi 1 -> 2 (03/10); trước đây tổ hợp này nổ NameError (`head_rows`).
"""
import datetime as dt
import os
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..")))
sys.path.insert(0, _HERE)

import openpyxl  # noqa: E402

import spec_extract as se  # noqa: E402

_loi = []


def _kiem(ten, thuc, mong):
    dat = thuc == mong
    print(f"  [{'OK ' if dat else 'SAI'}] {ten}: {thuc!r}")
    if not dat:
        _loi.append(ten)


def _ghi(ws, dong, cot, gia_tri):
    for j, v in enumerate(gia_tri):
        if v is not None:
            ws.cell(row=dong, column=cot + j, value=v)


def _sheet_xdv(wb, ten, thang, hdr, ngay_cot, co_tong_xdv, ten_dim1, so_chu=False):
    """Dựng 1 sheet kế hoạch ngày XDV, 2 ngày, mỗi ngày (Số lượng, Doanh thu)."""
    ws = wb.create_sheet(ten)
    ws.cell(row=hdr - 1, column=3, value="THÁNG")
    ws.cell(row=hdr - 1, column=4, value=thang)
    for k in range(2):
        ws.cell(row=hdr - 1, column=ngay_cot + 2 * k, value=dt.datetime(2026, thang, k + 1))
    _ghi(ws, hdr, 1, ["STT ", "Mã Code center" if co_tong_xdv else None, "XDV", ten_dim1])
    for k in range(2):
        _ghi(ws, hdr, ngay_cot + 2 * k, ["Số lượng", "Doanh thu"])
    r = hdr + 2                                     # chừa dòng phụ đề 'RO'
    khoi = ([("VINFAST_SDV", "Tổng XDV", 999)] if co_tong_xdv else []) + \
        [("OCP_XDV", "Ocean Park", 10), ("LB_XDV", "Long Biên", 20)]
    for ma, ten_x, dt_ngay in khoi:
        for nhan, he in (("Tổng", 1), ("N", 0.5)):
            dau = [None, ma, ten_x] if nhan == "Tổng" else [None, None, None]
            _ghi(ws, r, 1, dau + [nhan])
            for k in range(2):
                tien = dt_ngay * he * 1e9
                _ghi(ws, r, ngay_cot + 2 * k, [5 * he, f"  {tien:,.0f} " if so_chu else tien])
            r += 1
    ws.cell(row=r, column=1, value="Cộng")
    ws.cell(row=r, column=ngay_cot + 1, value=123456789e9)      # dòng Cộng: số rác, không được đọc
    return ws


def _spec_xdv():
    return {"report_type": "T", "nguon": {"sheet": {"moi_sheet_regex": "^kht\\d{1,2}ngay$"}},
            "chi_tu_ky": "2026-09", "giu_ngay_tuong_lai": True,
            "header": {"tim_o": "STT"},
            "dong_bat_dau_tim": {"cot": "B", "regex": "_XDV\\s*$"},
            "dong_ket_thuc_tim": {"cot": "A", "regex": "^\\s*Cộng"},
            "cot": {"cost_center": {"header": "XDV", "bat_buoc": False},
                    "dim1": {"header": ["Ma lenh", "HTTT"], "chuan_hoa": "cat"}},
            "ban_ghi": "moi_cot_ngay",
            "cot_ngay": {"dong_tuong_doi": -1, "tu": "F", "den": "Z", "lech_amount": 1,
                         "lech_amount2": 0, "he_so": 1e-09, "he_so_amount2": 1, "ky_tu_o": True},
            "loc": [{"cot": "dim1", "dieu_kien": "thuoc", "gia_tri": ["Tổng", "N", "W", "I", "C"]}]}


def _sheet_sr(wb, ten, thang, hdr):
    ws = wb.create_sheet(ten)
    ws.cell(row=hdr - 1, column=2, value=f"THÁNG {thang}")
    _ghi(ws, hdr, 1, ["Mã", "Kênh", None, None, "Cộng tháng", "Ngày 01", "Ngày 02"])
    r = hdr + 1
    for khoi, nhan, sl in (("A.", "A. Tổng sản lượng", 1), ("B.", "B.Tổng Doanh thu", 1e9)):
        _ghi(ws, r, 1, [khoi, nhan]); r += 1
        _ghi(ws, r, 1, ["B2C", "Kênh B2C", None, None, None, 3 * sl, 4 * sl]); r += 1
        _ghi(ws, r, 1, [None, "+ Vinfast Ocean Park", None, None, None, 1 * sl, 2 * sl]); r += 1


def _spec_sr(khoi):
    bat, het = (("^\\s*A\\.\\s*$", "^\\s*B\\.\\s*$") if khoi == "A"
                else ("^\\s*B\\.\\s*$", "^\\s*[C-Z]\\.\\s*$"))
    return {"report_type": "T", "nguon": {"sheet": {"moi_sheet_regex": "^kh.*ngay.*theosr$"}},
            "giu_ngay_tuong_lai": True, "ky_thang_tu_o": {"cot": "B", "toi_da": 30},
            "header": {"tim_o": "Mã"},
            "dong_bat_dau_tim": {"cot": "A", "regex": bat},
            "dong_ket_thuc_tim": {"cot": "A", "regex": het},
            "ngu_canh_dong": [
                {"khi": [{"cot": "A", "dieu_kien": "regex", "gia_tri": "^\\s*[AB]\\.\\s*$"}],
                 "xoa": ["dim2"]},
                {"khi": [{"cot": "A", "dieu_kien": "regex", "gia_tri": "^\\s*(B2C|B2B|GF)\\s*$"}],
                 "gan": {"dim2": {"cot": "A", "chuan_hoa": "hoa"}}}],
            "cot": {"payload.ten_dong": {"header": "Kênh", "chuan_hoa": "cat"}},
            "ban_ghi": "moi_cot_ngay",
            "cot_ngay": {"dong_tuong_doi": 0, "tu": "F", "den": "G", "he_so": 1.0},
            "loc": [{"cot": "dim2", "dieu_kien": "khac_rong"}]}


def _tong(recs, thang, **dk):
    return round(sum(r["amount"] for r in recs if r["ngay"][5:7] == f"{thang:02d}"
                     and all(r.get(k) == v for k, v in dk.items())), 6)


def main():
    tmp = tempfile.mkdtemp()

    print("XDV — hai bố cục trong cùng một file")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet_xdv(wb, "KHT8ngày", 8, 4, 9, True, "Ma lenh")
    _sheet_xdv(wb, "KHT9ngày", 9, 4, 9, True, "Ma lenh")
    _sheet_xdv(wb, "KHT10ngày", 10, 2, 8, False, "HTTT", so_chu=True)
    _sheet_xdv(wb, "KHT10ngày (2)", 10, 2, 8, False, "HTTT", so_chu=True)
    p = os.path.join(tmp, "2.XDV.OO.M.202608.Kehoachthang.xlsx")
    wb.save(p)
    recs, warn = se.extract_file(_spec_xdv(), p)
    se._dong_wb()
    ky = sorted({r["ngay"][:7] for r in recs})
    _kiem("chỉ T9 + T10 (T8 bị chi_tu_ky chặn)", ky, ["2026-09", "2026-10"])
    _kiem("T9: không lọt khối 'Tổng XDV' / dòng 'Cộng' (2 xưởng × 2 dòng × 2 ngày)",
          sum(r["ngay"][:7] == "2026-09" for r in recs), 8)
    _kiem("T9 Σ doanh thu dòng Tổng (tỷ)", _tong(recs, 9, dim1="Tổng"), 60.0)
    _kiem("T10: bản sao '(2)' không cộng đôi", sum(r["ngay"][:7] == "2026-10" for r in recs), 8)
    _kiem("T10 Σ doanh thu dòng Tổng — số ghi dạng chữ (tỷ)", _tong(recs, 10, dim1="Tổng"), 60.0)
    _kiem("T10 dim1 đọc từ cột 'HTTT'", sorted({r["dim1"] for r in recs if r["ngay"] >= "2026-10"}),
          ["N", "Tổng"])
    _kiem("số lượng vào amount2", sorted({r["amount2"] for r in recs}), [2.5, 5.0])

    print("SR — khối A./B. đổi dòng, nhãn tháng đổi ô")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet_sr(wb, "KHT9ngaytheoSR", 9, 8)
    _sheet_sr(wb, "KHngayT10theoSR", 10, 2)
    wb.create_sheet("KHngàytheokenh")              # sheet T8 bố cục khác — không được khớp regex
    p = os.path.join(tmp, "1.SR.OO.M.202608.Kehoachthang.xlsx")
    wb.save(p)
    a, _ = se.extract_file(_spec_sr("A"), p)
    b, _ = se.extract_file(_spec_sr("B"), p)
    se._dong_wb()
    _kiem("khối A: kỳ lấy từ nhãn tháng của từng sheet", sorted({r["ngay"] for r in a}),
          ["2026-09-01", "2026-09-02", "2026-10-01", "2026-10-02"])
    # dòng kênh là dòng TIÊU ĐỀ (chỉ gán dim2, không sinh bản ghi) -> chỉ còn dòng SR = 1 + 2
    _kiem("khối A T10 Σ (chỉ dòng SR)", _tong(a, 10), 3.0)
    _kiem("khối B không lẫn sản lượng của khối A", _tong(b, 9), 3e9)
    _kiem("dim2 = kênh", sorted({r["dim2"] for r in a + b}), ["B2C"])

    print("KHDT — header tự dò + cột theo tháng")
    for dong_hdr in (1, 2):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "KHDT"
        _ghi(ws, dong_hdr, 1, ["STT", "Mã số", "Khối", "Tổng năm", "Tháng 1", "Tháng 2"])
        _ghi(ws, dong_hdr + 1, 1, ["A", "GIÁ TRỊ"])
        _ghi(ws, dong_hdr + 2, 1, [None, "A200", "Doanh thu", 30, 10, 20])
        p = os.path.join(tmp, f"1.SR.OO.M.202608.Kehoachthang_h{dong_hdr}.xlsx")
        wb.save(p)
        spec = {"report_type": "T", "nguon": {"sheet": {"ten": "KHDT"}}, "header": {"tim_o": "Khối"},
                "cot": {"dim1": {"header": "Mã số"}}, "ban_ghi": "moi_cot_thang",
                "nam_tu_ten_file": {"regex": "\\.M\\.(\\d{4})"},
                "ky_thang_tu_ten_file": {"regex": "\\.M\\.(\\d{4})(\\d{2})"},
                "cot_thang": {"tu": "D", "den": "Z"},
                "loc": [{"cot": "dim1", "dieu_kien": "bang", "gia_tri": "A200"}]}
        recs, _ = se.extract_file(spec, p)
        se._dong_wb()
        _kiem(f"header dòng {dong_hdr}: 2 tháng, bỏ cột 'Tổng năm'",
              sorted((r["ngay"], r["amount"]) for r in recs),
              [("2026-01-31", 10.0), ("2026-02-28", 20.0)])

    print()
    if _loi:
        print(f"SAI {len(_loi)}: {_loi}")
        sys.exit(1)
    print("OK — tất cả đạt")


if __name__ == "__main__":
    main()

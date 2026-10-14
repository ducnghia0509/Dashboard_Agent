# -*- coding: utf-8 -*-
"""Test bot so khớp mapping — đọc tài liệu mapping + mã nguồn thật, DB TEST read-only (không ghi gì).
Chạy: .venv/bin/python scripts/test_mapping_check.py"""
import os
import sys
import tempfile

os.environ.setdefault("MAPPING_CACHE_DIR", tempfile.mkdtemp())
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from servers.common import mapping_check as mc  # noqa: E402

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
        return False
    except (ValueError, RuntimeError):
        return True


chk("_kd bỏ dấu/hoa/khoảng trắng lạ", mc._kd("  Tỷ\xa0lệ  Đã HOÀN thành ") == "ty le da hoan thanh")
d = mc.danh_muc("test")
chk("danh_muc: đủ 49 màn", d["so_man"] == 49)
chk("danh_muc: màn Trạm sạc chỉ có tài liệu md", {"tsac0", "tsac1", "tsac2"} <= set(d["man_chi_co_md"]))
chk("danh_muc: vhkd1 có tài liệu VHKD", next(x for x in d["man"] if x["man"] == "vhkd1")["tai_lieu_mapping"] == ["VHKD"])
chk("danh_muc: mỗi màn có builder", all(x["builder"] for x in d["man"]))
for man in ("vhkd0", "at1", "xdv0", "ts1", "htxt0", "duan0", "xt0", "revenue"):
    chk(f"màn {man} có logic", bool(mc.logic_man(man)["report_type_doc"]))
chk("màn lạ bị chặn", raises(mc.logic_man, "khong_co_man"))
chk("tìm chỉ tiêu có kết quả", mc.tim_chi_tieu("tỷ lệ hoàn thành doanh thu", man="vhkd0")["so_khop"] >= 1)
chk("tìm theo tài liệu lạ bị chặn", raises(mc.tim_chi_tieu, "doanh thu", tai_lieu="KHONGCO"))
chk("từ khoá quá ngắn bị chặn", raises(mc.tim_chi_tieu, "a"))
chk("không dấu = có dấu", mc.tim_chi_tieu("ty le hoan thanh doanh thu", man="vhkd0")["so_khop"] ==
    mc.tim_chi_tieu("tỷ lệ hoàn thành doanh thu", man="vhkd0")["so_khop"])
r = mc.doi_chieu("vhkd0", "kế hoạch doanh thu", env="test", period="2026-09")
chk("đối chiếu trả cờ nghi vấn là list", all(isinstance(x["co_nghi_van"], list) for x in r["doi_chieu"]))
chk("đối chiếu nêu builder + file:dòng", r["logic"]["builder"] and ":" in r["logic"]["file"])
chk("file Excel trích từ mô tả nguồn", mc._RE_FILE.findall("A:\\x\\1.SR.OO.M.202608.Kehoach.xlsx Sheet: CT") ==
    ["1.SR.OO.M.202608.Kehoach.xlsx"])
tv = mc.truy_vet("vhkd0", "tỷ lệ hoàn thành doanh thu", env="test", period="2026-09")
chk("truy vết: có trích nguyên văn mapping", tv["1_mapping_ghi"] and tv["1_mapping_ghi"][0]["nguyen_van"].get("chi_tieu"))
chk("truy vết: bắt được mã CT33 trong code", "CT33" in tv["2_code_dashboard"]["ma_chi_tieu_trong_code"])
chk("truy vết: có đoạn mã kèm file:dòng", tv["2_code_dashboard"]["doan_ma"] and ":" in tv["2_code_dashboard"]["doan_ma"][0]["dong"])
chk("truy vết: CT33 không bị coi là report_type", "CT33" not in tv["3_file_nguon_that"]["file_trong_db"])
chk("truy vết: có số dashboard đang hiện (CT33)", tv["4_so_dashboard_dang_hien"].get("so_muc_khop", 0) >= 1)
from servers.common import man_hien_thi as mh  # noqa: E402
chk("số hiển thị: CT33 có thucHien & nguon", any("thucHien" in x["gia_tri"] and x["gia_tri"].get("nguon")
    for x in mh.so_hien_thi("vhkd0", "CT33", "2026-09", "test")["so_hien_thi"]))
chk("số hiển thị: kỳ lạ bị chặn", raises(mh.so_hien_thi, "vhkd0", "CT33", "1999-01", "test"))
chk("số hiển thị: màn lạ lỗi rõ", raises(mh.so_hien_thi, "khong_co_man", "CT33", "2026-09", "test"))
chk("truy vết: luôn nêu việc CHƯA kiểm", len(tv["CHUA_KIEM"]) == 2)
chk("truy vết: chỉ tiêu lạ -> cờ KHONG_CO_DONG_MAPPING", any("KHONG_CO_DONG_MAPPING" in x for x in
    mc.truy_vet("vhkd0", "zzzz không tồn tại", env="test")["co_nghi_van"]))
print(f"{ok} PASS, {len(loi)} FAIL")
for t in loi:
    print("  FAIL:", t)
sys.exit(1 if loi else 0)

# -*- coding: utf-8 -*-
"""Báo động FILE ĐANG BỊ ẨN sau khi trích xuất (`spec_extract.file_dang_an` + `run`).

Chạy: python scripts/test_canh_bao_file_an.py
      (có DATABASE_URL -> thêm 1 phép kiểm CHỈ ĐỌC trên DB thật)

Vì sao có test này (15/10/2026): script dọn cộng đôi 10/10 ẩn theo tên
`HUNGTHINH::B.5.HT.M2026.08.BC cap nhat van hanh.xlsx`; 11/10 máy gửi phát hành lại đúng tên cũ ->
`cron_htxt_daily` mỗi sáng log `written=124/82` như thành công, htxt0 trống SL xe bán + tiến trình
hợp đồng. Test khoá:
  · file bị ẩn mà có dòng ghi -> có cảnh báo chứa "CẢNH BÁO" (panel bắt theo chữ này) + khoá
    `file_dang_an`, ở cả --write lẫn dry-run;
  · file không bị ẩn / ghi 0 dòng -> IM, không báo động giả;
  · `source_file` hỏi bảng phải KHỚP đúng cái `_ghi` ghi (kể cả hậu tố `source_file_hau_to`);
  · không đọc được DB -> rỗng, không chặn đường nạp;
  · cron_htxt_daily rút đúng dòng cảnh báo từ JSON stdout.
"""
import json
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..")))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..", "cron")))
sys.path.insert(0, _HERE)

import spec_extract as se  # noqa: E402
import cron_htxt_daily as ch  # noqa: E402

loi = []


def chk(ten, dk):
    print(f"  [{'OK ' if dk else 'SAI'}] {ten}")
    if not dk:
        loi.append(ten)


PATH = "/x/received_reports/HUNGTHINH/baocaocapnhatvanhanhxetai/B.5.HT.M2026.08.BC cap nhat van hanh.xlsx"
SF = "HUNGTHINH::B.5.HT.M2026.08.BC cap nhat van hanh.xlsx"
RECS = [{"ngay": "2026-10-01", "amount": 1}] * 3

# Cô lập `run` khỏi file thật và DB thật.
se.extract_file = lambda spec, path: (list(RECS), [])
se._bo_ngay_tuong_lai = lambda spec, recs: (recs, [])
se._kiem_thang_ten_file = lambda spec, path, recs: None
se._kiem_cong_doi = lambda spec, recs: None
_hoi = []


def gia_lap(an, written=3):
    se._ghi = lambda spec, path, recs: {"written": written}

    def _an(ds):
        _hoi.append(list(ds))
        return {x for x in ds if x in an}
    se.file_dang_an = _an


SPEC = {"id": "htxt_giao", "report_type": "HTXT_GIAO"}

print("1. File bị ẩn, --write ghi 3 dòng -> BÁO ĐỘNG")
gia_lap({SF})
r = se.run(SPEC, PATH, write=True)
cb = " ".join(r.get("canh_bao") or [])
chk("có khoá file_dang_an đúng source_file", r.get("file_dang_an") == SF)
chk("cảnh báo có chữ 'CẢNH BÁO' + report_type + số dòng",
    "CẢNH BÁO" in cb and "HTXT_GIAO" in cb and "3 dòng vừa ghi" in cb)

print("2. Dry-run, file bị ẩn -> vẫn báo (số dòng SẼ bị che)")
r = se.run(SPEC, PATH, write=False)
chk("dry-run báo 'sẽ ghi'", "3 dòng sẽ ghi" in " ".join(r.get("canh_bao") or []))

print("3. File KHÔNG bị ẩn -> im")
gia_lap(set())
r = se.run(SPEC, PATH, write=True)
chk("không có cảnh báo, không có khoá", not r.get("canh_bao") and "file_dang_an" not in r)

print("4. Ghi 0 dòng (dù file bị ẩn) -> im, và không hỏi DB")
gia_lap({SF}, written=0)
_hoi.clear()
r = se.run(SPEC, PATH, write=True)
chk("không cảnh báo khi written=0", not r.get("canh_bao") and not _hoi)

print("5. Hậu tố source_file_hau_to -> hỏi đúng tên `_ghi` ghi")
gia_lap({SF + "#KHT10"})
_hoi.clear()
r = se.run({**SPEC, "source_file_hau_to": "KHT10"}, PATH, write=True)
chk("hỏi bảng với '<FOLDER>::<file>#KHT10'", _hoi == [[SF + "#KHT10"]])
chk("báo động bắn cho tên có hậu tố", r.get("file_dang_an") == SF + "#KHT10")

print("6. Không có DB -> rỗng, không nổ")
import importlib  # noqa: E402
se2 = importlib.reload(se)
_db = se2.DB_URL
se2.DB_URL = None
chk("DB_URL rỗng -> set()", se2.file_dang_an([SF]) == set())
se2.DB_URL = "postgresql://khong:co@127.0.0.1:1/khongco"
chk("DB không kết nối được -> set()", se2.file_dang_an([SF]) == set())
se2.DB_URL = _db

print("7. cron_htxt_daily rút cảnh báo từ stdout JSON")
out = json.dumps([{"file": "a.xlsx", "dong": 124, "written": 124, "file_dang_an": SF,
                   "canh_bao": ["Ô LỖI EXCEL x", "CẢNH BÁO FILE ĐANG BỊ ẨN (hidden_files): ..."]},
                  {"file": "b.xlsx", "dong": 5, "canh_bao": ["CẢNH BÁO khác"]},
                  {"_bo_qua_anh_chup_cu": []}], ensure_ascii=False)
chk("chỉ lấy dòng hidden_files của file bị ẩn",
    ch.canh_bao_file_an(out) == ["CẢNH BÁO FILE ĐANG BỊ ẨN (hidden_files): ..."])
chk("stdout hỏng -> []", ch.canh_bao_file_an("không phải json") == [])

if os.environ.get("DATABASE_URL"):
    print("8. DB thật (CHỈ ĐỌC): file HTXT đang ẩn 15/10/2026")
    an = se2.file_dang_an([SF, "HUNGTHINH::khong_co_file_nay.xlsx"])
    print(f"     -> {sorted(an)}")
    chk("trả đúng các tên có trong hidden_files (không lẫn tên lạ)",
        "HUNGTHINH::khong_co_file_nay.xlsx" not in an)

print(f"\n{'HỎNG ' + str(len(loi)) if loi else 'ĐẠT'}")
sys.exit(1 if loi else 0)

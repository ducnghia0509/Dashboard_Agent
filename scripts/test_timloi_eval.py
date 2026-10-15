# -*- coding: utf-8 -*-
"""EVAL + test đơn vị cho MCP `dashboard_timloi` (bot tìm lỗi / nghiệm thu), 15/10/2026 — hợp đồng đợt 2+3 mục C.

Chạy:  .venv/bin/python scripts/test_timloi_eval.py            (eval đọc PROD 5434, chỉ đọc)
       .venv/bin/python scripts/test_timloi_eval.py --env test  (eval trên TEST 5435)
Không cần container/LLM: gọi THẲNG hàm Python sau tool. Mọi kết nối DB read-only; file chỉ đọc.

PHẦN A — EVAL bắt buộc (3 ca thật ngày 15/10 + 3 ca tool gộp của giám sát):
  1. htxt1 nhập kho T7/2026 = 20 xe / 1.720.000 USD (03_LoHang: Tình trạng 'Đã nhận', Ngày hàng về tháng 7, bỏ phụ tùng)
  2. htxt0 SL xe bán T8/2026 = 7 (raw_rows HTXT_GIAO lọc payload.tinh_trang = Bán; builder kpi 'sl')
  3. Thuế Dự án T01/2026: CĐKT mã 313 cuối kỳ = 91.018.726 đồng ↔ dashboard thue.phaiNop.cuoi = 0,091018726 tỷ
  4. tl_truy_vet(htxt1, donNhap.tongXe.nhanSl, 2026-07) = 20
  5. tl_kiem_nhanh_nap('HTXT_GIAO') KHÔNG còn HIDDEN (user đã bỏ ẩn 15/10) — và giả lập hidden thì HIDDEN bật
  6. tl_so_qua_ky(thue, phaiNop.cuoi, T01..T03, khối Dự án) = 0.091018726 / 0.232918225 / 0.877532704
PHẦN B — an toàn: chặn đường ngoài thư mục/symlink, chặn tham số kiểu SQL injection, kết nối DB chỉ đọc.
PHẦN C — quy ước đọc Excel (ngày gõ tay/serial, dấu cách/hoa thường, Decimal) trên workbook tổng hợp trong thư mục tạm.
PHẦN D — ví dụ JSON trong 2 SKILL hợp lệ theo schema phiếu v1.
"""
import json
import os
import re
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(__file__), "..")))
from servers.common import so_lieu_check as sc  # noqa: E402
from servers.common import tim_loi as tl  # noqa: E402
from servers.common import tim_loi_excel as tx  # noqa: E402

ENV = "test" if "--env" in sys.argv and sys.argv[sys.argv.index("--env") + 1] == "test" else "prod"
ok, loi, bang, gio = 0, [], [], {}


def chk(ten, cond):
    global ok
    if cond:
        ok += 1
    else:
        loi.append(ten)


def ev(ca, ky_vong, ra):
    khop = ra is not None and abs(float(ra) - float(ky_vong)) <= 1e-9 * max(1.0, abs(float(ky_vong)))
    bang.append((ca, ky_vong, ra, khop))
    chk(f"EVAL {ca}: kỳ vọng {ky_vong}, tool ra {ra}", khop)


def do_gio(ten, f):
    t = time.time()
    r = f()
    gio[ten] = round(time.time() - t, 3)
    return r


try:
    sc._ro(ENV, "SELECT 1 AS x")
except Exception as e:  # noqa: BLE001
    if ENV == "prod":
        print(f"[!] không đọc được PROD ({e}) — chạy eval trên TEST")
        ENV = "test"
    else:
        raise

# ═════════════════════════════ A. EVAL ═════════════════════════════
F_HD = "HUNGTHINH/theodoidonhangtheohopdong/Theo_doi_don_hang_theo_hop_dong_08.2026.xlsx"
LOC_NK = [{"cot": "Tình trạng", "toan_tu": "=", "gia_tri": "Đã nhận"},
          {"cot": "Ngày hàng về", "toan_tu": "thang", "gia_tri": "2026-07"},
          {"cot": "Loại hàng", "toan_tu": "!=", "gia_tri": "phụ tùng"}]
r = do_gio("tl_tong_hop_excel", lambda: tx.tong_hop_excel(F_HD, "03_LoHang", 4, LOC_NK, ["Số lượng", "Giá trị lô (NT)"]))
ev("1a htxt1 T7 nhập kho — số xe (file)", 20, r["tong"]["Số lượng"])
ev("1b htxt1 T7 nhập kho — giá trị USD (file)", 1720000, r["tong"]["Giá trị lô (NT)"])
chk("1 dòng góp phần có số dòng Excel", r["dong_gop_phan"] and all(isinstance(d["dong"], int) for d in r["dong_gop_phan"]))
chk("1 tổng chính xác dạng chuỗi Decimal", r["tong_chinh_xac"]["Giá trị lô (NT)"] == "1720000")
r2 = tx.tong_hop_excel(F_HD, "03_LoHang", 4, LOC_NK[:2] + [{"cot": "H", "toan_tu": "khong_chua", "gia_tri": "PHỤ TÙNG"}],
                       ["I", "O"])
chk("1 cùng số khi gọi cột bằng chữ Excel + khong_chua", r2["tong"] == {"I": 20, "O": 1720000})

r = tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01", "2026-08-31",
                loc=[{"cot": "payload.tinh_trang", "toan_tu": "=", "gia_tri": "Bán"}])
ev("2a htxt0 T8 SL xe bán — raw_rows HTXT_GIAO (Bán)", 7, r["tong_bo_file_dang_an"]["amount2"])
r = do_gio("tl_so_dashboard", lambda: tl.so_dashboard(ENV, "htxt0", "2026-08", "month", key="kpi.sl", bo_cache=True))
ev("2b htxt0 T8 kpi sl (dashboard)", 7, r["so"])
ev("2c htxt0 T8 kpi.sl.thucHien (dashboard)", 7, tl.so_dashboard(ENV, "htxt0", "2026-08", key="kpi.sl.thucHien")["so"])

F_BC = "DUAN/baocaotaichinhrieng/B.4.TC.TCKT.M.202601.Baocaotaichinhrieng.xlsx"
r = do_gio("tl_doc_excel", lambda: tx.doc_excel(F_BC, "CĐKT", header_dong=8, tim="313"))
d313 = [d for d in r["dong"] if str(d.get("Mã số")).strip() == "313"]
ev("3a Thuế Dự án T01 CĐKT mã 313 cuối kỳ (file, đồng)", 91018726, d313[0]["Số cuối kỳ"] if d313 else None)
r = tl.so_dashboard(ENV, "thue", "2026-01", "month", khoi="Khối KD Dự án", key="phaiNop.cuoi")
ev("3b thue T01 phaiNop.cuoi Khối Dự án (dashboard, tỷ)", 0.091018726, r["so"])

r = do_gio("tl_truy_vet", lambda: tl.truy_vet("htxt1", "donNhap.tongXe.nhanSl", "2026-07", ENV, so_dung=20, bo_cache=True))
ev("4 tl_truy_vet htxt1 donNhap.tongXe.nhanSl T7", 20, r["1_dashboard"]["so"])
chk("4 truy vết có raw_rows HTXT_NCC_LO amount2 = 20",
    any(t["report_type"] == "HTXT_NCC_LO" and t["amount2"] == 20 for t in r["3_raw_rows_theo_report_type"]))
chk("4 gợi ý tầng lệch = khong", r["goi_y_tang_lech"]["tang"] == "khong")

r = do_gio("tl_kiem_nhanh_nap", lambda: tl.kiem_nhanh_nap("HTXT_GIAO", env=ENV))
chk("5a HTXT_GIAO không còn cảnh báo HIDDEN", "HIDDEN" not in r["ma_canh_bao"])
bang.append(("5a tl_kiem_nhanh_nap HTXT_GIAO — HIDDEN", "không có", ",".join(r["ma_canh_bao"]) or "(rỗng)",
             "HIDDEN" not in r["ma_canh_bao"]))
sf = [x["source_file"] for x in r["theo_report_type"]["HTXT_GIAO"]]
_an_that = tl._an
tl._an = lambda env: set(sf[:1])                          # giả lập user ẩn file — KHÔNG ghi DB
try:
    rg = tl.kiem_nhanh_nap("HTXT_GIAO", env=ENV)
    chk("5b giả lập hidden -> HIDDEN bật", "HIDDEN" in rg["ma_canh_bao"])
    rr = tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01", "2026-08-31")
    chk("5c giả lập hidden -> raw_rows cờ dang_an + tổng bỏ ẩn = 0",
        rr["file_dang_an_co_dong"] == sf[:1] and rr["tong_bo_file_dang_an"]["so_dong"] == 0 and rr["tong"]["so_dong"] == 7)
    tv = tl.truy_vet("htxt0", "kpi.sl", "2026-08", ENV, so_dung=7)
    chk("5d truy vết đọc cờ ẩn ở raw_rows", any(f["dang_an"] for t in tv["3_raw_rows_theo_report_type"]
                                                for f in t.get("source_file", [])))
finally:
    tl._an = _an_that
bang.append(("5b giả lập hidden — HIDDEN", "có", ",".join(rg["ma_canh_bao"]), "HIDDEN" in rg["ma_canh_bao"]))

r = do_gio("tl_so_qua_ky", lambda: tl.so_qua_ky(ENV, "thue", "phaiNop.cuoi", ["2026-01", "2026-02", "2026-03"],
                                                  khoi="Khối KD Dự án", bo_cache=True))
for k, v in zip(["2026-01", "2026-02", "2026-03"], [0.091018726, 0.232918225, 0.877532704]):
    ev(f"6 tl_so_qua_ky thue phaiNop.cuoi {k}", v, next((d["so"] for d in r["bang"] if d["ky"] == k), None))
r = do_gio("tl_so_qua_ky(ca_hai, cache)", lambda: tl.so_qua_ky("ca_hai", "thue", "phaiNop.cuoi", ["2026-01"],
                                                                 khoi="Khối KD Dự án"))
chk("6 ca_hai trả cả prod lẫn test", {d["env"] for d in r["bang"]} == {"prod", "test"})
r = do_gio("tl_ho_so_man", lambda: tl.ho_so_man("htxt1", bo_cache=True))
chk("hồ sơ htxt1: builder + HTXT_NCC_LO có spec htxt_ncc_lo",
    r["builder"]["ham"] == "build_htxt1" and [s["id"] for s in r["spec_theo_report_type"]["HTXT_NCC_LO"]] == ["htxt_ncc_lo"])
chk("hồ sơ htxt1: có dòng mapping + trường FE donNhap", r["so_dong_mapping"] > 0 and "donNhap" in r["fe"]["truong_d_tren_trang"])
do_gio("tl_ho_so_man(cache)", lambda: tl.ho_so_man("htxt1"))
m = tl.doc_mapping(tai_lieu="HTXT", dong=58)
chk("doc_mapping HTXT dòng 58 nguyên văn có 'Đã nhận'", "Đã nhận" in json.dumps(m["dong"], ensure_ascii=False))
chk("doc_mapping báo 2 bản cùng tên (bản copy)", len(m["ban_cung_ten"]) == 2)
f = tl.tim_file_nguon("hungthinh", "theodoidonhangtheohopdong")
chk("tim_file_nguon: thấy file đơn hàng + giờ lưu về 252", f["file"] and f["file"][0]["luu_ve_252"])
c = tl.doc_code(ham="_where")
chk("doc_code _where ở app/metrics/repository.py", c["tim_thay"] and c["tim_thay"][0]["file"] == "app/metrics/repository.py")
c = tl.doc_code(man="htxt1")
chk("doc_code man htxt1 -> thân build_htxt1", c["builder"] == "build_htxt1" and "def build_htxt1" in c["tim_thay"][0]["ma"])

# ═════════════════════════════ B. AN TOÀN ═════════════════════════════
def chan(ten, f, loai=ValueError):
    try:
        f()
        chk(f"chặn: {ten}", False)
    except loai:
        chk(f"chặn: {ten}", True)
    except Exception as e:  # noqa: BLE001
        chk(f"chặn: {ten} (lỗi sai loại {type(e).__name__}: {e})", False)


for xau in ("../etc/passwd.xlsx", "/etc/passwd", "/etc/x.xlsx", "HUNGTHINH/../../../../etc/x.xlsx",
            "/home/itadmin/AI_Dashboard_QT/Dashboard_Agent/x.xlsx", "a\x00.xlsx", "", None,
            F_HD.replace(".xlsx", ".json")):
    chan(f"đường file {xau!r}", lambda: tx.duong_an_toan(xau))
tmp = tempfile.mkdtemp(prefix="timloi_")
goc_cu = list(tx.GOC_CHO_PHEP)
try:
    ngoai = os.path.join(tmp, "ngoai")
    trong = os.path.join(tmp, "goc", "DV", "tm")
    os.makedirs(ngoai)
    os.makedirs(trong)
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["tiêu đề"])
    ws.append(["Ngày", "Tình trạng", "SL", "Tiền"])
    rows = [("05/07/2026", "  Đã   NHẬN ", 1, 0.1), (46213, "đã nhận", 2, 0.2), ("2026-07-31", "Đã nhận", "1.000", "#REF!"),
            ("31/7/26", "Đã nhận", 3, 1e-9), ("ngày lạ", "Đã nhận", 5, 5), ("01/08/2026", "Đã nhận", 7, 7),
            (None, "Đã nhận", 9, 9), ("05/07/2026", "Da nhan", 11, 11)]
    for x in rows:
        ws.append(list(x))
    p_that = os.path.join(ngoai, "bi_mat.xlsx")
    wb.save(p_that)
    wb.save(os.path.join(trong, "that.xlsx"))
    os.symlink(p_that, os.path.join(trong, "link_ra_ngoai.xlsx"))
    os.symlink(os.path.join(trong, "that.xlsx"), os.path.join(trong, "link_trong.xlsx"))
    os.symlink(ngoai, os.path.join(tmp, "goc", "DV", "tm_link"))
    tx.GOC_CHO_PHEP[:] = [os.path.join(tmp, "goc")]
    chan("symlink trỏ ra ngoài gốc", lambda: tx.duong_an_toan("DV/tm/link_ra_ngoai.xlsx"))
    chan("symlink trong gốc", lambda: tx.duong_an_toan("DV/tm/link_trong.xlsx"))
    chan("thư mục symlink", lambda: tx.duong_an_toan("DV/tm_link/bi_mat.xlsx"))
    chan("file ngoài gốc (tuyệt đối)", lambda: tx.duong_an_toan(p_that))
    chk("file thật trong gốc đọc được", tx.duong_an_toan("DV/tm/that.xlsx").endswith("that.xlsx"))

    # ═════════════════ C. QUY ƯỚC ĐỌC EXCEL ═════════════════
    r = tx.tong_hop_excel("DV/tm/that.xlsx", "S", 2, [{"cot": "tình TRẠNG", "toan_tu": "=", "gia_tri": "đã nhận"},
                                                      {"cot": "Ngày", "toan_tu": "thang", "gia_tri": "2026-07"}],
                          ["SL", "Tiền"])
    # khớp: 05/07 (dd/mm, dấu cách thừa+hoa), serial 46213 (=2026-07-10), 2026-07-31, 31/7/26 -> 4 dòng
    chk(f"C ngày gõ tay/serial/ISO/2 số năm + chuỗi không phân biệt hoa thường: 4 dòng (ra {r['so_dong_khop']})",
        r["so_dong_khop"] == 4)
    chk(f"C SL cộng cả ô chữ '1.000' kiểu VN: 1+2+1000+3 = 1006 (ra {r['tong']['SL']})", r["tong"]["SL"] == 1006)
    chk(f"C Decimal không sai số float: 0.1+0.2+1e-9 (ra {r['tong_chinh_xac']['Tiền']})",
        r["tong_chinh_xac"]["Tiền"] == "0.300000001")
    chk("C ô '#REF!' đếm vào o_khong_phai_so", r["o_khong_phai_so"].get("Tiền", {}).get("so_o_loi") == 1)
    chk("C ô 'ngày lạ' đếm vào o_ngay_khong_doc_duoc", r["o_ngay_khong_doc_duoc"].get("Ngày", {}).get("so_o_loi") == 1)
    chk("C giữ dấu: 'Da nhan' KHÔNG bằng 'đã nhận'", all(d["dong"] != 10 for d in r["dong_gop_phan"]))
    r = tx.tong_hop_excel("DV/tm/that.xlsx", "S", 2, [{"cot": "Ngày", "toan_tu": "khoang_ngay",
                                                      "gia_tri": ["2026-07-31", "2026-08-01"]}], ["SL"], nhom_theo="Ngày")
    chk(f"C khoang_ngay gồm 2 đầu: 1000+3+7 (ra {r['tong']['SL']})", r["tong"]["SL"] == 1010)
    chan("toan_tu lạ", lambda: tx.tong_hop_excel("DV/tm/that.xlsx", "S", 2, [{"cot": "SL", "toan_tu": "LIKE", "gia_tri": 1}], ["SL"]))
    chan("cột không có", lambda: tx.tong_hop_excel("DV/tm/that.xlsx", "S", 2, None, ["Không có"]))
finally:
    tx.GOC_CHO_PHEP[:] = goc_cu
    shutil.rmtree(tmp, ignore_errors=True)
chk("C doc_ngay serial 46213 = 2026-07-10 (gốc 1899-12-30)", str(tx.doc_ngay(46213)) == "2026-07-10")
chk("C doc_ngay '1/2/2026' = 2026-02-01 (ngày trước tháng)", str(tx.doc_ngay("1/2/2026")) == "2026-02-01")
chk("C doc_ngay '30/02/2026' = None", tx.doc_ngay("30/02/2026") is None)
chk("C doc_so '1.000' = 1000, '1.5' = 1.5, '0.100' = 0.1, '1.720.000,5' = 1720000.5",
    [str(tx.doc_so(x)) for x in ("1.000", "1.5", "0.100", "1.720.000,5")] == ["1000", "1.5", "0.100", "1720000.5"])

# SQL injection / tham số lạ: mọi tham số đi qua kiểm hoặc placeholder.
INJ = "HTXT_GIAO' OR '1'='1"
chan("report_type injection", lambda: tl.raw_rows(ENV, INJ, "2026-08-01", "2026-08-31"))
chan("report_type chữ thường", lambda: tl.raw_rows(ENV, "htxt_giao", "2026-08-01", "2026-08-31"))
chan("ngày injection", lambda: tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01' OR '1'='1", "2026-08-31"))
chan("ngày không tồn tại", lambda: tl.raw_rows(ENV, "HTXT_GIAO", "2026-02-30", "2026-03-01"))
chan("nhom_theo injection", lambda: tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01", "2026-08-31", nhom_theo="dim1; DROP TABLE raw_rows"))
chan("loc cot injection", lambda: tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01", "2026-08-31",
                                              loc=[{"cot": "amount) OR (1=1", "toan_tu": "=", "gia_tri": 1}]))
chan("loc payload khoá lạ", lambda: tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01", "2026-08-31",
                                                loc=[{"cot": "payload.a'--", "toan_tu": "=", "gia_tri": 1}]))
r = tl.raw_rows(ENV, "HTXT_GIAO", "2026-08-01", "2026-08-31",
                loc=[{"cot": "dim1", "toan_tu": "=", "gia_tri": "x' OR '1'='1"}])
chk("giá trị lọc kiểu injection chỉ là chuỗi so sánh -> 0 dòng", r["tong"]["so_dong"] == 0)
chan("env lạ", lambda: tl.raw_rows("staging", "HTXT_GIAO", "2026-08-01", "2026-08-31"))
chan("env ca_hai ở tool đơn env", lambda: tl.so_dashboard("ca_hai", "htxt0", "2026-08"))
chan("kiem_nhanh_nap đối tượng lạ", lambda: tl.kiem_nhanh_nap("x'; DROP TABLE raw_rows; --", env=ENV))
chan("kiem_nhanh_nap folder có ..", lambda: tl.kiem_nhanh_nap("../../etc", env=ENV))
chan("so_dashboard kỳ lạ", lambda: tl.so_dashboard(ENV, "htxt0", "2026-08; rm -rf /"))
chan("so_dashboard màn lạ", lambda: tl.so_dashboard(ENV, "khong co man nao nhu the nay xyz", "2026-08"))
chan("doc_code tên hàm lạ", lambda: tl.doc_code(ham="x(); import os"))
chan("tim_file_nguon đơn vị là đường", lambda: tl.tim_file_nguon("../.."))
chan("tim_file_nguon thư mục là đường", lambda: tl.tim_file_nguon("HUNGTHINH", "../../etc"))
chan("doc_mapping file ngoài Tài liệu", lambda: tl.doc_mapping(file="/etc/passwd.xlsx"))
chan("so_qua_ky quá 24 kỳ", lambda: tl.so_qua_ky(ENV, "thue", "phaiNop.cuoi", ["2026-01"] * 25))

# Read-only: kết nối của tool bị Postgres từ chối ghi; tiến trình builder nhận PGOPTIONS read-only.
for e in ("test", "prod"):
    try:
        with sc._ket_noi(e) as cn:
            cn.execute("CREATE TEMP TABLE _timloi_x(a int)")
        chk(f"kết nối {e} CHỈ ĐỌC", False)
    except Exception as ex:  # noqa: BLE001
        chk(f"kết nối {e} CHỈ ĐỌC", "read-only" in str(ex).lower())
for e in ("test", "prod"):
    ev_ = tl._env_api(e)
    chk(f"builder {e}: PGOPTIONS read-only + đúng cổng", "default_transaction_read_only=on" in ev_["PGOPTIONS"]
        and f":{tl._PORT[e]}/" in ev_["DATABASE_URL"])
chk("lay_key: phần tử danh sách theo 'ma'", tl.lay_key({"kpi": [{"ma": "dt", "thucHien": 1}, {"ma": "sl", "thucHien": 7}]},
                                                       "kpi.sl.thucHien")[0] == 7)
chk("lay_key: báo lỗi key thiếu", tl.lay_key({"a": {}}, "a.b")[2] is not None)

# ═════════════════════════════ D. JSON MẪU TRONG SKILL ═════════════════════════════
try:
    import jsonschema
    _root = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
    sp = next(p for p in (os.path.join(_root, "agents", "phieu_loi.schema.json"),
                          "/home/itadmin/AI_Dashboard_QT/wt/phieu_loi.schema.json") if os.path.exists(p))
    with open(sp, encoding="utf-8") as fh:
        schema = json.load(fh)
    for agent, khoi in (("tim_loi", "chan_doan"), ("nghiem_thu", "nghiem_thu")):
        with open(os.path.join(_root, "agents", agent, "SKILL.md"), encoding="utf-8") as fh:
            khoi_json = re.findall(r"```json\n(.*?)\n```", fh.read(), re.S)
        sub = dict(schema["properties"][khoi])
        sub["$defs"] = schema["$defs"]
        sub["additionalProperties"] = False                 # SKILL cấm khoá ngoài schema — test chặt hơn schema
        try:
            jsonschema.validate(json.loads(khoi_json[0]), sub)
            chk(f"SKILL {agent}: ví dụ {khoi} hợp lệ schema v1", len(khoi_json) == 1)
        except Exception as ex:  # noqa: BLE001
            chk(f"SKILL {agent}: ví dụ {khoi} hợp lệ schema v1 ({ex})", False)
except ImportError:
    chk("jsonschema có trong venv", False)

print(f"\nEVAL (env={ENV}):")
print("| ca | số kỳ vọng | số tool ra | khớp |")
for ca, kv, ra, kh in bang:
    print(f"| {ca} | {kv} | {ra} | {'CÓ' if kh else 'KHÔNG'} |")
print("\nThời gian (giây):", json.dumps(gio, ensure_ascii=False))
print(f"\n{ok} PASS, {len(loi)} FAIL")
for t in loi:
    print("  FAIL:", t)
sys.exit(1 if loi else 0)

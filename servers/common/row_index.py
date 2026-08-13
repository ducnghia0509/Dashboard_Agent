# -*- coding: utf-8 -*-
"""CHỈ MỤC NHÃN DÒNG — "trong sheet có dòng gì", thứ mà source_catalog KHÔNG có.

VÌ SAO CÓ FILE NÀY (2026-08-13): catalog chỉ biết tên file / tên sheet / tên cột, nên
`catalog_search(query="doanh thu")` LUÔN rỗng — không sheet nào tên vậy. Agent muốn tìm một chỉ
tiêu thì chỉ còn cách mở file rồi đổ 200 dòng thô ra dò, nhân với 11 đơn vị. Đo được hậu quả:
payload p90 = 34 KB/lượt gọi, 55 lần compaction trên 15 phiên.

Chỉ mục này quét cột nhãn (A–C) của MỌI sheet trong catalog và lưu "nhãn này nằm ở file/sheet/dòng
nào". Nhờ đó `tim_chi_tieu()` trả về vị trí chính xác trong vài mili-giây, rồi `source_inspect` chỉ
đọc đúng vùng đó.

ĐÂY LÀ CHỈ MỤC DẪN ĐƯỜNG, KHÔNG PHẢI KHO SỐ LIỆU: chỉ ghi nhãn + toạ độ, KHÔNG ghi giá trị. Xoá
lúc nào cũng được, dựng lại từ Excel ~10 phút. Ràng buộc "qa chỉ đọc Excel" giữ nguyên.

Dựng lại TĂNG DẦN theo mtime — file mới đẩy về mới phải quét, nên nguồn báo cáo mới tự vào chỉ mục
mà không ai phải khai gì.

Chạy: python -m servers.common.row_index [--rebuild]
"""
import os
import sqlite3
import sys
import unicodedata

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
DB_PATH = os.path.join(_AGENT_ROOT, "memory", "row_label_index.sqlite")

# Trần quét. Sheet BCTC có thể khai dimension tới cả triệu dòng phantom (xem source_catalog:142)
# nên PHẢI có trần, nếu không một file làm treo cả vòng dựng.
MAX_DONG = 600
MAX_COT_NHAN = 3          # nhãn chỉ tiêu nằm ở cột A–C ở mọi mẫu báo cáo đang có
MIN_DAI_NHAN = 3

_SCHEMA = """
CREATE TABLE IF NOT EXISTS nhan (
  file_key TEXT, file TEXT, path TEXT, company TEXT, report_type TEXT,
  year INTEGER, month INTEGER, ky TEXT, period_type TEXT,
  sheet TEXT, dong INTEGER, ma_dong TEXT, nhan TEXT, nhan_dac TEXT
);
CREATE INDEX IF NOT EXISTS ix_nhan_rt   ON nhan(report_type, ky);
CREATE INDEX IF NOT EXISTS ix_nhan_file ON nhan(file_key);
CREATE INDEX IF NOT EXISTS ix_nhan_ma   ON nhan(ma_dong);
CREATE TABLE IF NOT EXISTS quet (file_key TEXT PRIMARY KEY, path TEXT, mtime REAL, so_nhan INTEGER);
CREATE VIRTUAL TABLE IF NOT EXISTS nhan_fts USING fts5(nhan_dac, content='nhan', content_rowid='rowid');
"""


def _dac(s) -> str:
    """Bỏ dấu + thường hoá. Câu hỏi người dùng gõ không dấu/sai chính tả vẫn phải khớp."""
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").strip().lower()


def _connect(path: str = None) -> sqlite3.Connection:
    con = sqlite3.connect(path or DB_PATH)
    con.row_factory = sqlite3.Row
    con.executescript(_SCHEMA)
    return con


def _la_ma_dong(v) -> bool:
    """Mã dòng = token ngắn dạng '01' / 'A100' / 'T100' / 'B300' / '60' đứng ở cột nhãn."""
    s = str(v or "").strip()
    return bool(s) and len(s) <= 6 and any(c.isdigit() for c in s) and " " not in s


def _quet_file(entry: dict) -> list:
    """Trả danh sách bản ghi nhãn của 1 file. Lỗi mở file -> ném lên cho người gọi ghi nhận,
    KHÔNG nuốt: file hỏng mà im lặng bỏ qua thì sau này trông y hệt 'đơn vị không phát sinh'."""
    from . import be_bridge as bb

    path = entry["path"]
    wb = bb.fast_load_workbook(path, data_only=True, read_only=True)
    ra = []
    try:
        for ws in wb.worksheets:
            for i, row in enumerate(ws.iter_rows(min_col=1, max_col=MAX_COT_NHAN,
                                                 max_row=MAX_DONG, values_only=True), start=1):
                nhan = ma = None
                for c in row:
                    if c is None:
                        continue
                    s = str(c).strip()
                    if not s:
                        continue
                    if ma is None and _la_ma_dong(s):
                        ma = s
                        continue
                    if nhan is None and len(s) >= MIN_DAI_NHAN and not s.replace(".", "").replace(",", "").isdigit():
                        nhan = s
                if not nhan:
                    continue
                ra.append((entry["file_key"], entry["file"], path, entry.get("company"),
                           entry.get("report_type"), entry.get("year"), entry.get("month"),
                           entry.get("ky"), entry.get("period_type"),
                           ws.title, i, ma, nhan[:300], _dac(nhan)[:300]))
    finally:
        wb.close()
    return ra


def dung_lai(rebuild: bool = False, log=print) -> dict:
    """Dựng/cập nhật chỉ mục. Mặc định TĂNG DẦN (bỏ qua file mtime không đổi).

    Ghi vào file .tmp rồi os.replace — đổi chỗ nguyên tử. Không được để `tim_chi_tieu` đọc trúng
    chỉ mục đang dựng dở rồi báo 'không tìm thấy' như thể dữ liệu không tồn tại.
    """
    from . import source_catalog as sc

    entries = sc.search()
    for e in entries:
        e["file_key"] = sc._file_key(e["path"])

    tmp = DB_PATH + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)
    if os.path.exists(DB_PATH) and not rebuild:
        import shutil
        shutil.copy2(DB_PATH, tmp)
    con = _connect(tmp)

    da_quet = {r["file_key"]: r["mtime"] for r in con.execute("SELECT file_key, mtime FROM quet")}
    moi = bo_qua = 0
    loi = []
    tong_nhan = 0
    for n, e in enumerate(entries, 1):
        fk = e["file_key"]
        try:
            mtime = os.path.getmtime(e["path"])
        except OSError:
            loi.append({"file": e["file"], "loi": "không mở được đường dẫn"})
            continue
        if not rebuild and da_quet.get(fk) == mtime:
            bo_qua += 1
            continue
        try:
            ra = _quet_file(e)
        except Exception as ex:                                    # noqa: BLE001
            loi.append({"file": e["file"], "loi": f"{type(ex).__name__}: {ex}"[:200]})
            continue
        con.execute("DELETE FROM nhan WHERE file_key=?", (fk,))
        con.executemany("INSERT INTO nhan VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", ra)
        con.execute("INSERT OR REPLACE INTO quet VALUES (?,?,?,?)", (fk, e["path"], mtime, len(ra)))
        moi += 1
        tong_nhan += len(ra)
        if n % 25 == 0:
            log(f"  ... {n}/{len(entries)} file, {tong_nhan} nhãn mới")
    con.execute("INSERT INTO nhan_fts(nhan_fts) VALUES('rebuild')")
    con.commit()
    tong = con.execute("SELECT COUNT(*) c FROM nhan").fetchone()["c"]
    con.close()
    os.replace(tmp, DB_PATH)
    return {"file_quet_moi": moi, "file_bo_qua": bo_qua, "tong_nhan": tong, "loi": loi}


def san_sang() -> dict:
    """Chỉ mục đã dựng chưa. Chưa dựng KHÁC HẲN 'không có dữ liệu' — phải phân biệt được, nếu
    không agent sẽ trả lời 'không tìm thấy' cho một hệ thống chỉ đơn giản là chưa index."""
    if not os.path.exists(DB_PATH):
        return {"san_sang": False, "ly_do": "chỉ mục chưa được dựng lần nào",
                "cach_sua": "chạy `python -m servers.common.row_index --rebuild`"}
    con = _connect()
    try:
        n = con.execute("SELECT COUNT(*) c FROM nhan").fetchone()["c"]
        f = con.execute("SELECT COUNT(*) c FROM quet").fetchone()["c"]
    finally:
        con.close()
    if n == 0:
        return {"san_sang": False, "ly_do": "chỉ mục rỗng", "cach_sua": "chạy lại với --rebuild"}
    return {"san_sang": True, "so_nhan": n, "so_file": f}


def tim(ten: str, ky: str = None, year=None, month=None, company: str = None,
        report_type: str = None, gioi_han: int = 40, nhom: str = None) -> dict:
    """Tra vị trí một chỉ tiêu: trả [{file, sheet, dong, ma_dong, nhan}] — KHÔNG mở file, KHÔNG
    trả giá trị. Bước tiếp theo là `source_inspect` đọc đúng vùng đó."""
    tt = san_sang()
    if not tt.get("san_sang"):
        return {"ket_qua": [], "count": 0, "chi_muc": tt}

    tu = [t for t in _dac(ten).split() if t]
    con = _connect()
    try:
        sql = ["SELECT file, path, company, report_type, ky, year, month, sheet, dong, ma_dong, "
               "nhan, nhan_dac FROM nhan WHERE 1=1"]
        args = []
        for t in tu:
            sql.append("AND nhan_dac LIKE ?")
            args.append(f"%{t}%")
        if ky:
            sql.append("AND ky=?")
            args.append(ky)
        if year not in (None, ""):
            sql.append("AND year=?")
            args.append(int(year))
        if month not in (None, ""):
            sql.append("AND month=?")
            args.append(int(month))
        if company:
            sql.append("AND company=?")
            args.append(company)
        if report_type:
            sql.append("AND report_type=?")
            args.append(report_type)
        if nhom:
            # Lọc theo THƯ MỤC NGUỒN (SRVF/XDV/DUAN…) — chiều người dùng thật sự hỏi. Lấy danh
            # sách file từ catalog thay vì thêm cột vào chỉ mục: khỏi phải quét lại 370 file.
            from . import source_catalog as sc
            ten_file = [e["file"] for e in sc.search(nhom=nhom)]
            if not ten_file:
                return {"ket_qua": [], "count": 0,
                        "canh_bao": f"Không có file nào thuộc nhóm nguồn '{nhom}'."}
            sql.append("AND file IN (" + ",".join("?" * len(ten_file)) + ")")
            args += ten_file
        sql.append("ORDER BY report_type, company, sheet, dong LIMIT 4000")
        rows = [dict(r) for r in con.execute(" ".join(sql), args)]
    finally:
        con.close()

    # Xếp hạng: dòng TIÊU ĐỀ chỉ tiêu (nhãn ngắn, khớp gần trọn, có mã dòng) đứng trước dòng chi
    # tiết. Sắp theo file/dòng như SQL trả về thì kết quả bị một công ty có nhiều sheet chiếm hết.
    dich = " ".join(tu)

    def _diem(r):
        n = r["nhan_dac"] if "nhan_dac" in r.keys() else _dac(r["nhan"])
        return (0 if n.startswith(dich) else 1, 0 if r.get("ma_dong") else 1, len(n))

    for r in rows:
        r["_d"] = _diem(r)
    rows.sort(key=lambda r: r.pop("_d"))

    # Tóm tắt độ phủ: agent cần biết NGAY là có bao nhiêu đơn vị/kỳ có chỉ tiêu này, để phát hiện
    # thiếu đơn vị trước khi cộng — thay vì đếm tay danh sách rồi kết luận nhầm là đủ.
    from collections import Counter
    tom_tat = {
        "so_vi_tri": len(rows),
        "theo_cong_ty": dict(Counter(r.get("company") or "(không rõ)" for r in rows)),
        "theo_report_type": dict(Counter(r.get("report_type") or "(không rõ)" for r in rows)),
        "theo_ky": dict(Counter(r.get("ky") or "(không rõ kỳ)" for r in rows)),
        "so_file": len({r["file"] for r in rows}),
    }
    bi_cat = len(rows) > gioi_han
    out = {"ket_qua": rows[:gioi_han], "count": min(len(rows), gioi_han), "tom_tat": tom_tat}
    if bi_cat:
        # Cắt danh sách thì PHẢI ghi nhãn. Cắt ngầm làm sai cả phép đếm lẫn phép cộng ở đầu ra.
        out["canh_bao"] = (f"Khớp {len(rows)} vị trí, đang hiển thị {gioi_han} vị trí phù hợp nhất. "
                           f"`tom_tat` phản ánh TOÀN BỘ {len(rows)} vị trí — dùng nó để kiểm đủ/thiếu "
                           f"đơn vị, đừng đếm theo danh sách đã cắt. Thu hẹp bằng ky/company/report_type.")
    return out


if __name__ == "__main__":
    rebuild = "--rebuild" in sys.argv
    sys.path.insert(0, _AGENT_ROOT)
    print(f"Dựng chỉ mục nhãn dòng ({'toàn bộ' if rebuild else 'tăng dần'}) -> {DB_PATH}")
    kq = dung_lai(rebuild=rebuild)
    print(f"\nQuét mới {kq['file_quet_moi']} file, bỏ qua {kq['file_bo_qua']} file "
          f"(mtime không đổi). Tổng {kq['tong_nhan']} nhãn.")
    if kq["loi"]:
        print(f"\n{len(kq['loi'])} file LỖI:")
        for x in kq["loi"]:
            print("  -", x["file"], "|", x["loi"])

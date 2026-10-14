# -*- coding: utf-8 -*-
"""BỘ KIỂM SỐ LIỆU TẤT ĐỊNH cho bot `ai_check_bot` (14/10/2026).

QUY TẮC CỐT LÕI (xem KIENTRUC_AI_DULIEU_QUYETDINH.md): LLM KHÔNG BAO GIỜ tự tính ra một con số. Mọi
phép cộng/so sánh nằm ở ĐÂY, bằng SQL cố định có tham số; bot chỉ nhận kết quả có cấu trúc rồi diễn
giải. Không có tool nào nhận SQL tự do.

AN TOÀN DỮ LIỆU: mọi kết nối mở ở chế độ `default_transaction_read_only=on` + `statement_timeout`.
Chọn DB bằng tham số `env` = prod (5434) | test (5435) — KHÔNG đọc `DATABASE_URL_RO` cũ (trỏ 5433, đã chết).

BA PHÉP KIỂM (khớp ba câu người dùng hay hỏi):
  1. `thang_vs_ngay`   BC tháng có bằng Σ BC ngày không — từng chỉ tiêu, chỉ ra dòng lệch.
  2. `chi_tiet_vs_tong` số hiển thị (PTHU_D…) có bằng Σ dòng chi tiết nguồn không — chỉ ra dòng bị bỏ
                       (vd XDV 12/10: 45,3678 vs 45,3832, chênh đúng 1 dòng R/O không tên khách).
  4. `file_ve_luc`      MỘT file nguồn: sinh ra ở máy nguồn lúc nào, về 252 lần đầu/lần cuối lúc nào, đã vào DB
                       những gì, cron nhắc tới nó ở đâu.
  3. `nguon_thieu_cham` file ngày/tháng của một đơn vị: file nào đã về, ngày nào CHƯA có file, ngày nào
                       có file mà DB chưa có số, cron lượt cuối nói gì.
Đơn vị tiền: các cột `amount` trong raw_rows của khối P&L/công nợ đã là TỶ ĐỒNG.
"""
import datetime
import glob
import json
import os
import re
from urllib.parse import unquote, urlparse

import psycopg
from psycopg.rows import dict_row

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
# Đường dẫn TUYỆT ĐỐI (như cron/pull_nguon.py): worktree của Dashboard_Agent nằm ngoài cây
# AI_Dashboard_QT nên đường tương đối `..` sẽ trỏ sai.
ROOT = os.environ.get("TC_ROOT") or "/home/itadmin/AI_Dashboard_QT"
CONNECT_ROOT = os.path.join(ROOT, "Connect_VPS")
RECEIVED_DIR = os.environ.get("RECEIVED_DIR") or os.path.join(CONNECT_ROOT, "received_reports")
LOGS_DIR = os.path.join(ROOT, "AI_coding", "logs")

_PORT = {"prod": 5434, "test": 5435}
# Chỉ ngưỡng TUYỆT ĐỐI (đơn vị tỷ): 0,2 triệu đồng. KHÔNG có ngưỡng tương đối — ngưỡng 0,5% đã làm lọt
# đúng ca cần bắt (XDV 12/10 chênh 15,4 triệu trên 45,38 tỷ = 0,03%; An Taxi doanh thu chênh 15,6 triệu).
NGUONG_TUYET_DOI = 2e-4

# Cặp (báo cáo THÁNG, báo cáo NGÀY) cùng nội dung. So theo `dim1` vì `dim2` khác nhau giữa hai bên
# (CHIPHI tháng tách 'CP nhân sự'/'Hoạt động khác' ở dim2, ngày để trống).
CAP_THANG_NGAY = {"PNLT": "PNLT_D", "DTHU": "DTHU_D", "CHIPHI": "CHIPHI_D", "HQKD": "HQKD_D"}

# Chi tiết nguồn -> số đã gom hiển thị, và khoá payload mang tên đối tượng mà bộ gom dùng để nhóm.
CAP_CHI_TIET = {
    "XDV_CN_RO_D": {"tong": "PTHU_D", "khoa_ten": "ten_doi_tuong", "ma": "ma_doi_tuong"},
    "VHKD_PTHU_HD": {"tong": "PTHU_D", "khoa_ten": "ten_khach", "ma": "ma_khach"},
}


def _ket_noi(env: str):
    if env not in _PORT:
        raise ValueError("env phải là 'prod' hoặc 'test'")
    dsn = os.environ.get("DATABASE_URL") or ""
    # Mật khẩu trong DSN là URL-ENCODED (`tc_%24production`): phải giải mã, không thì gửi nguyên chuỗi
    # `%24` lên Postgres và bị "password authentication failed" — lỗi chỉ lộ khi chạy dưới service có
    # nạp .env (serve_http), còn chạy tay khi chưa đặt DATABASE_URL thì dùng mặc định nên không thấy.
    mat_khau = unquote(urlparse(dsn).password or "") if dsn else ""
    mat_khau = mat_khau or "tc_$production"
    return psycopg.connect(
        host="localhost", port=_PORT[env], dbname="tc_dashboard", user="tc", password=mat_khau,
        autocommit=True, row_factory=dict_row, prepare_threshold=None,
        options="-c default_transaction_read_only=on -c statement_timeout=20000")


def _ro(env, sql, params=()):
    with _ket_noi(env) as c:
        return c.execute(sql, params).fetchall()


def _f(x):
    return float(x) if x is not None else None


def don_vi_co_trong_db(env="prod") -> list:
    """Tiền tố nguồn (trước '::' của source_file) — dùng làm tham số `don_vi`."""
    rows = _ro(env, "SELECT DISTINCT split_part(source_file,'::',1) AS dv FROM raw_rows "
                    "WHERE source_file LIKE '%%::%%' ORDER BY 1")
    return [r["dv"] for r in rows]


def _chuan_don_vi(env, don_vi):
    co = don_vi_co_trong_db(env)
    for d in co:
        if d.upper() == str(don_vi or "").strip().upper():
            return d, co
    raise ValueError(f"không có đơn vị '{don_vi}' (nguồn có: {', '.join(co)})")


def _lech(a, b):
    """-> (chênh b−a, % theo |a|, khớp?)."""
    if a is None or b is None:
        return None, None, False
    d = b - a
    pct = (d / abs(a)) if abs(a) > 1e-12 else None
    khop = abs(d) <= NGUONG_TUYET_DOI
    return round(d, 6), (None if pct is None else round(pct * 100, 2)), khop


# ───────────────────────────── 1. THÁNG vs Σ NGÀY ─────────────────────────────
def thang_vs_ngay(don_vi: str, period: str, env: str = "prod") -> dict:
    """So từng chỉ tiêu BC THÁNG với Σ BC NGÀY của `period` (YYYY-MM) cho một đơn vị."""
    dv, _ = _chuan_don_vi(env, don_vi)
    if not re.fullmatch(r"\d{4}-\d{2}", period or ""):
        raise ValueError("period phải dạng YYYY-MM")
    pref = dv + "::%"
    ket_qua, ghi_chu = [], []
    for rt_t, rt_n in CAP_THANG_NGAY.items():
        t = {r["dim1"] or "": r for r in _ro(
            env, "SELECT dim1, SUM(amount) AS s, COUNT(*) AS n, MAX(source_file) AS sf FROM raw_rows "
                 "WHERE period_month=%s AND report_type=%s AND source_file LIKE %s GROUP BY dim1",
            (period, rt_t, pref))}
        n = {r["dim1"] or "": r for r in _ro(
            env, "SELECT dim1, SUM(amount) AS s, COUNT(*) AS n, COUNT(DISTINCT ngay) AS ng "
                 "FROM raw_rows WHERE period_month=%s AND report_type=%s AND source_file LIKE %s "
                 "GROUP BY dim1", (period, rt_n, pref))}
        if not t and not n:
            continue
        for k in sorted(set(t) | set(n)):
            a, b = _f(t.get(k, {}).get("s")), _f(n.get(k, {}).get("s"))
            d, pct, khop = _lech(a, b)
            ket_qua.append({"bao_cao": rt_t, "chi_tieu": k, "thang": a, "tong_ngay": b,
                            "chenh_ngay_tru_thang": d, "chenh_pct": pct,
                            "ket_luan": ("KHỚP" if khop else
                                         "THIẾU Ở THÁNG" if a is None else
                                         "THIẾU Ở NGÀY" if b is None else "LỆCH"),
                            "so_ngay_co_so": n.get(k, {}).get("ng")})
    if not ket_qua:
        raise ValueError(f"{dv} {period}: không có cả BC tháng lẫn BC ngày trong DB {env}")
    lech = [r for r in ket_qua if r["ket_luan"] != "KHỚP"]
    # Ngày trong kỳ có số / không có số ở BC ngày.
    ngay_co = [r["ngay"] for r in _ro(
        env, "SELECT DISTINCT ngay FROM raw_rows WHERE period_month=%s AND source_file LIKE %s "
             "AND report_type IN ('PNLT_D','HQKD_D') AND ngay IS NOT NULL ORDER BY 1", (period, pref))]
    y, m = int(period[:4]), int(period[5:])
    cuoi = (datetime.date(y + (m == 12), m % 12 + 1, 1) - datetime.timedelta(days=1)).day
    het = [f"{period}-{d:02d}" for d in range(1, cuoi + 1) if f"{period}-{d:02d}" not in ngay_co]
    if het:
        ghi_chu.append(f"BC ngày thiếu {len(het)} ngày có P&L: {', '.join(x[8:] for x in het)}")
    return {"don_vi": dv, "period": period, "env": env, "don_vi_tien": "tỷ đồng",
            "so_chi_tieu": len(ket_qua), "so_lech": len(lech),
            "tong_hop": ("Tháng = Σ ngày ở mọi chỉ tiêu" if not lech else
                         f"{len(lech)}/{len(ket_qua)} chỉ tiêu lệch"),
            "chi_tiet": ket_qua, "ngay_thieu": het, "ghi_chu": ghi_chu,
            "goi_y_nguyen_nhan": ([
                "Bút toán chốt tháng (phân bổ, khấu hao, lương, lãi vay, thu nhập/chi phí khác) thường "
                "chỉ có ở BC tháng mà BC ngày 'tạm tính' chưa có — kiểm với kế toán, KHÔNG tự kết luận.",
                "Mở nguồn so lại: file tháng và file ngày có cùng bản phát hành/cùng ngày chốt chưa."]
                if lech else [])}


# ───────────────────────────── 2. CHI TIẾT vs TỔNG ─────────────────────────────
def chi_tiet_vs_tong(don_vi: str, ngay: str, rt_chi_tiet: str = "XDV_CN_RO_D", env: str = "prod") -> dict:
    """Σ dòng chi tiết nguồn (`rt_chi_tiet`) có bằng số gom hiển thị (PTHU_D…) cùng ngày không.

    Dòng bị bỏ thường là dòng KHÔNG CÓ TÊN ĐỐI TƯỢNG: bộ gom nhóm theo tên nên dòng rỗng tên rơi mất.
    """
    if rt_chi_tiet not in CAP_CHI_TIET:
        raise ValueError(f"rt_chi_tiet phải là một trong {sorted(CAP_CHI_TIET)}")
    cfg = CAP_CHI_TIET[rt_chi_tiet]
    dv, _ = _chuan_don_vi(env, don_vi)
    pref = dv + "::%"
    ct = _ro(env, f"SELECT amount, dim1, dim2, payload::json->>%s AS ten, payload::json->>%s AS ma, "
                  "source_file FROM raw_rows WHERE report_type=%s AND ngay=%s AND source_file LIKE %s",
             (cfg["khoa_ten"], cfg["ma"], rt_chi_tiet, ngay, pref))
    tong = _ro(env, "SELECT SUM(amount) AS s, COUNT(*) AS n FROM raw_rows WHERE report_type=%s AND ngay=%s "
                    "AND (source_file LIKE %s OR source_file LIKE %s)",
               (cfg["tong"], ngay, pref, "TEST_" + dv + "::%"))[0]
    if not ct:
        raise ValueError(f"{dv} {ngay}: không có dòng chi tiết {rt_chi_tiet}")
    s_ct = sum(_f(r["amount"]) or 0 for r in ct)
    co_ten = [r for r in ct if (r["ten"] or "").strip()]
    khong_ten = [r for r in ct if not (r["ten"] or "").strip()]
    s_ten = sum(_f(r["amount"]) or 0 for r in co_ten)
    s_tong = _f(tong["s"])
    d, pct, khop = _lech(s_ct, s_tong)
    return {"don_vi": dv, "ngay": ngay, "env": env, "don_vi_tien": "tỷ đồng",
            "chi_tiet_nguon": {"report_type": rt_chi_tiet, "so_dong": len(ct), "tong": round(s_ct, 6),
                               "file": sorted({r["source_file"] for r in ct})},
            "so_hien_thi": {"report_type": cfg["tong"], "so_dong": tong["n"],
                            "tong": None if s_tong is None else round(s_tong, 6)},
            "chenh_hien_thi_tru_chi_tiet": d, "chenh_pct": pct,
            "ket_luan": "KHỚP" if khop else "LỆCH",
            "dong_khong_ten_doi_tuong": {
                "so_dong": len(khong_ten), "tong": round(sum(_f(r["amount"]) or 0 for r in khong_ten), 6),
                "mau": [{"amount": _f(r["amount"]), "dim1": r["dim1"], "ma": r["ma"]}
                        for r in sorted(khong_ten, key=lambda r: -abs(_f(r["amount"]) or 0))[:5]]},
            "tong_dong_co_ten": round(s_ten, 6),
            "giai_thich": (
                None if khop else
                "Chênh bằng đúng tổng các dòng không có tên đối tượng." if
                abs((d or 0) + sum(_f(r["amount"]) or 0 for r in khong_ten)) <= NGUONG_TUYET_DOI else
                "Chênh KHÔNG hoàn toàn do dòng thiếu tên — cần soi thêm (cấn trừ, nhóm có số dư ròng 0, "
                "hoặc bản nạp cũ).")}


# ───────────────────────────── 3. NGUỒN THIẾU / CHẬM ─────────────────────────────
_RE_NGAY_TEN = re.compile(r"\.[DM]\.(\d{4})(\d{2})(\d{2})(?!\d)")


def _ngay_tu_ten(ten):
    m = _RE_NGAY_TEN.search(ten)
    if not m:
        return None
    try:
        return datetime.date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    except ValueError:
        return None


def nguon_thieu_cham(don_vi: str, period: str, env: str = "prod", thu_muc: str = "baocaohqkdngay") -> dict:
    """File ngày của đơn vị: đã về / chưa về / có file mà DB chưa có số + cron lượt cuối."""
    dv, _ = _chuan_don_vi(env, don_vi)
    goc = dv[5:] if dv.startswith("TEST_") else dv
    folder = os.path.join(RECEIVED_DIR, goc, thu_muc)
    if not os.path.isdir(folder):
        raise ValueError(f"không có thư mục nguồn {goc}/{thu_muc}")
    y, m = int(period[:4]), int(period[5:])
    dau = datetime.date(y, m, 1)
    cuoi = datetime.date(y + (m == 12), m % 12 + 1, 1) - datetime.timedelta(days=1)
    hom_nay = datetime.date.today()
    cuoi_xet = min(cuoi, hom_nay - datetime.timedelta(days=1))
    file_ngay = {}
    for p in glob.glob(os.path.join(folder, "*.xls*")):
        ten = os.path.basename(p)
        ng = _ngay_tu_ten(ten)
        if ng and ng[:7] == period:
            file_ngay.setdefault(ng, []).append(
                {"file": ten, "ve_luc": datetime.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(
                    timespec="minutes")})
    co_so = {r["ngay"] for r in _ro(
        env, "SELECT DISTINCT ngay FROM raw_rows WHERE period_month=%s AND source_file LIKE %s "
             "AND ngay IS NOT NULL AND report_type IN ('BS_D','PNLT_D','HQKD_D','PTHU_D','HH_D')",
        (period, dv + "::%"))}
    chua_file, co_file_chua_so = [], []
    d = dau
    while d <= cuoi_xet:
        iso = d.isoformat()
        if iso not in file_ngay:
            chua_file.append(iso)
        elif iso not in co_so:
            co_file_chua_so.append(iso)
        d += datetime.timedelta(days=1)
    # Log cron: dòng 'kiem [<ĐƠN VỊ>] <kỳ>' gần nhất.
    cron = None
    log = os.path.join(LOGS_DIR, "cron_hqkdngay_daily_prod.log" if env == "prod" else "cron_hqkdngay_daily.log")
    if os.path.exists(log):
        with open(log, encoding="utf-8", errors="replace") as fh:
            for dong in fh.readlines()[-4000:][::-1]:
                if f"kiem [{goc}] {period}" in dong:
                    cron = dong.strip()[:300]
                    break
    return {"don_vi": dv, "period": period, "env": env, "thu_muc": thu_muc,
            "so_file_ngay_ve": len(file_ngay), "ngay_da_co_file": sorted(file_ngay),
            "ngay_CHUA_co_file": chua_file, "ngay_co_file_ma_DB_chua_co_so": co_file_chua_so,
            "xet_den_ngay": cuoi_xet.isoformat(), "cron_luot_cuoi": cron,
            "ket_luan": ("Nguồn đủ và DB đã có số" if not chua_file and not co_file_chua_so else
                         ("Có ngày CHƯA CÓ FILE nguồn — việc của bên phát hành file (kế toán), cron không "
                          "kéo được cái chưa tồn tại." if chua_file else "") +
                         (" Có file nhưng DB chưa có số — kiểm log nạp (lệch tên sheet/bố cục)." if co_file_chua_so
                          else "")),
            "luu_y": "File ngày N của một số đơn vị (An Taxi…) chỉ chứa P&L tới N−1; ngày cuối tháng "
                     "có thể nằm ở file đầu tháng sau."}


# ───────────────────────────── 4. MỘT FILE KÉO LÚC NÀO ─────────────────────────────
def _doc_json(p):
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def file_ve_luc(ten_file: str, env: str = "prod") -> dict:
    """Lịch sử của MỘT file nguồn theo TÊN: sinh ở máy nguồn · về 252 · vào DB · cron.

    Mốc thời gian (đều là GIỜ VIỆT NAM, đọc từ sidecar `.json` cạnh file):
      `tao_o_may_nguon` / `sua_o_may_nguon` = office_created_at / modified_at — lúc kế toán/Cyber xuất file;
      `ve_252_lan_dau`  = created_at — lần đầu receiver nhận;  `ve_252_lan_cuoi` = saved_at — lần kéo gần nhất
      (file giữ nguyên tên bị GHI ĐÈ mỗi lần kéo nên saved_at là lần ghi đè cuối, không phải lần đầu).
    KHÔNG có mốc "đã nạp vào DB lúc nào": raw_rows không lưu giờ nạp — chỉ nêu được số dòng/ngày đã có.
    """
    ten = os.path.basename(str(ten_file or "").strip())
    if not re.fullmatch(r"[\w .()\-]+\.(xlsx|xlsm|xlsb|xls)", ten, flags=re.IGNORECASE):
        raise ValueError("ten_file phải là tên file Excel (không kèm đường dẫn)")
    thay = []
    for p in glob.glob(os.path.join(RECEIVED_DIR, "*", "*", "*")):
        if os.path.basename(p).lower() == ten.lower():
            sc = _doc_json(os.path.splitext(p)[0] + ".json")
            thay.append({
                "thu_muc": os.path.relpath(os.path.dirname(p), RECEIVED_DIR),
                "tao_o_may_nguon": sc.get("office_created_at"), "sua_o_may_nguon": sc.get("modified_at"),
                "ve_252_lan_dau": sc.get("created_at"), "ve_252_lan_cuoi": sc.get("saved_at"),
                "nguoi_tao": sc.get("created_by"), "nguoi_sua_cuoi": sc.get("last_modified_by"),
                "duong_dan_o_may_nguon": sc.get("original_source_path"),
                "dung_luong_byte": os.path.getsize(p)})
    if not thay:
        raise ValueError(f"không thấy file '{ten}' trong received_reports (kiểm tên file, hoặc file chưa từng về)")
    db = _ro(env, "SELECT report_type, COUNT(*) AS so_dong, MIN(ngay) AS tu_ngay, MAX(ngay) AS den_ngay, "
                  "MAX(source_file) AS khoa_nguon FROM raw_rows WHERE source_file LIKE %s "
                  "GROUP BY report_type ORDER BY 1", ("%::" + ten,))
    # Số suy ra từ file NGÀY (deriver) ghi source_file = '<ĐƠN VỊ>::<thư mục>::YYYY-MM', KHÔNG mang tên file → tra theo
    # (đơn vị, thư mục, ngày trong tên file). Không có nhánh này thì file đã nạp đủ vẫn hiện "chưa có trong DB".
    theo_ngay = []
    m_ngay = re.search(r"\.D\.(\d{4})(\d{2})(\d{2})\.", ten)
    if not db and m_ngay and thay:
        dv, _, tm = thay[0]["thu_muc"].partition("/")
        ngay = f"{m_ngay.group(1)}-{m_ngay.group(2)}-{m_ngay.group(3)}"
        theo_ngay = _ro(env, "SELECT report_type, COUNT(*) AS so_dong, MAX(source_file) AS khoa_nguon FROM raw_rows "
                             "WHERE ngay=%s AND source_file LIKE %s GROUP BY report_type ORDER BY 1",
                        (ngay, f"{dv.upper()}::{tm.lower()}::%"))
    nhat_ky = []
    for log in sorted(glob.glob(os.path.join(LOGS_DIR, "cron_*.log"))):
        if "dryrun" in log:
            continue
        try:
            with open(log, encoding="utf-8", errors="replace") as fh:
                for dong in fh.readlines()[-60000:]:
                    if ten in dong:
                        nhat_ky.append({"log": os.path.basename(log), "dong": dong.strip()[:240]})
        except OSError:
            continue
    return {"ten_file": ten, "env": env, "gio": "giờ Việt Nam", "ban_ghi": thay,
            "trong_db": [{"report_type": r["report_type"], "so_dong": r["so_dong"],
                          "tu_ngay": r["tu_ngay"], "den_ngay": r["den_ngay"], "khoa_nguon": r["khoa_nguon"]}
                         for r in db],
            "trong_db_theo_ngay_cua_file": [{"report_type": r["report_type"], "so_dong": r["so_dong"],
                                            "khoa_nguon": r["khoa_nguon"]} for r in theo_ngay],
            "cron_nhac_toi_file": nhat_ky[-5:],
            "luu_y": "`trong_db` rỗng KHÔNG có nghĩa chưa nạp: số suy ra từ file ngày nằm ở `trong_db_theo_ngay_cua_file` (khoá nguồn "
                     "theo đơn vị+thư mục+tháng, không mang tên file; KHÔNG chứng minh được dòng nào do đúng file này). "
                     "Không có mốc 'nạp vào DB lúc nào' (raw_rows không lưu giờ nạp). Nếu cần chứng minh thứ tự "
                     "kéo→nạp, so `ve_252_lan_cuoi` với dòng cron `autofill` cùng file."}


def to_json(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2, default=str)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP BÁO CÁO NGÀY (HQKD theo ngày, 12 đơn vị) tự động — 3 lượt/ngày.

BA LƯỢT/NGÀY: 05:00 · 16:45 · 17:15 giờ VN (đổi 29/08/2026 theo yêu cầu user; trước đó một lượt 17:00 VN, và test/prod ĐỀU ở 10:00 UTC nên đang tranh khoá per-file).
Lượt PROD chạy TRƯỚC TEST đúng 10 phút (04:50 · 16:35 · 17:05 VN; đổi 29/08/2026 theo yêu cầu
user, trước đó prod chạy SAU test 5 phút). KHOẢNG CÁCH là bắt buộc, chiều nào cũng được miễn khác
phút: cùng phút thì hai lượt tranh khoá per-file (servers/common/filelock.py), lượt sau bị
'skipped_lock' và MẤT HẲN một lượt nạp. Đánh đổi: test hết vai 'chim báo bão' — nguồn hỏng nay
prod vấp TRƯỚC, nên khi một khối im số thì đọc log PROD trước.
MỐC CRONTAB VIẾT THEO UTC (máy TZ=Etc/UTC, cron Ubuntu bỏ qua CRON_TZ khi TÍNH LỊCH — man 5
crontab, LIMITATIONS): TEST 09:45 · 10:15 · 22:00 UTC, PROD 09:35 · 10:05 · 21:50 UTC. Lượt sáng
(05:00 VN test / 04:50 VN prod) nằm ở 22:00 và 21:50 UTC HÔM TRƯỚC.


Mirror cron_thuchi_daily.py (cùng thư mục) — xem đó để hiểu khung chung (vì sao POST
/request-file thay vì sync_orchestrator pull, vì sao chờ mtime/saved_at thay vì sleep cứng, vì sao
mốc crontab phải viết theo UTC). Khác THUCHI ở 2 điểm:

1. MỘT report_type ("baocaohqkdngay") nhưng 12 ĐƠN VỊ độc lập (SRVF/XANHVINHPHUC/HTXXANHTUYENQUANG/
   HTXXANHVINHPHUC/ANTAXI/ANKHACHSAN/GLOBALAI/TRAMSAC/DUAN/HO/HUNGTHINH/XDV — khớp _UNITS trong
   Dashboard_Agent/scripts/derive_hqkd_ngay.py) — mỗi đơn vị 1 file riêng, KHÔNG có phụ thuộc
   chéo giữa các file (khác THUCHI: chain VAY đọc ngược file ngân hàng) -> xử lý độc lập, thứ tự
   không quan trọng, thiếu vài đơn vị không chặn các đơn vị còn lại.
2. KHÔNG cần xoá tay trước khi nạp lại: mỗi file báo cáo ngày là 1 workbook LUỸ KẾ CẢ THÁNG (mỗi
   ngày đã qua = 1 sheet/cột), và derive_hqkd_ngay.derive() (dispatch tự động qua agent_cli.py
   autofill -> is_daily_report() gate) tự làm "DELETE FROM raw_rows WHERE source_file=... rồi nạp
   lại TOÀN BỘ các ngày đọc được trong CHÍNH LẦN GỌI ĐÓ" (xem derive_hqkd_ngay.py dòng ~1187) —
   idempotent sẵn, gọi lại bao nhiêu lần cũng ra đúng 1 bản mới nhất, không cộng dồn/trùng ngày.
   receiver_server ghi file mới đè "wb" nên cũng không cần xoá file cũ trên đĩa trước khi kéo lại.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_hqkdngay_daily.py [--dry-run] [--env test|prod]

VỊ TRÍ: script nằm trong repo `Dashboard_Agent` (chuyển vào 2026-08-12; trước đó ở `AI_coding/` —
thư mục ngoài mọi git repo nên không có version control) nhưng LOG + ARTIFACT vẫn ghi vào
`AI_coding/logs/` — CỐ Ý: agent giám sát trên openclaw và hằng số `_LOGS` của `source_bridge.py`
(nút "Chạy ngay"/"Tắt" ở tab Nguồn dữ liệu) đều đang trỏ vào đó. Đổi chỗ log là phải sửa đồng thời
3 nơi, không đáng.
"""
import argparse
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))   # cron_status.py cùng thư mục
import cron_status                                               # noqa: E402

VN = timezone(timedelta(hours=7))
ROOT = "/home/itadmin/AI_Dashboard_QT"
CONNECT = f"{ROOT}/Connect_VPS"
AVAILABLE_META = f"{CONNECT}/available_metadata.json"
RECEIVED_DIR = f"{CONNECT}/received_reports"
AGENT = f"{ROOT}/Dashboard_Agent"
AGENT_PY = f"{AGENT}/.venv/bin/python"
RECEIVER_URL = "http://127.0.0.1:8090"

ENVIRONMENTS = {
    "test": {
        "database_url": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
        "log_file": f"{ROOT}/AI_coding/logs/cron_hqkdngay_daily.log",
        "disabled_flag": f"{ROOT}/AI_coding/logs/cron_hqkdngay_daily.disabled",
        "verify_api_dir": f"{ROOT}/AI_coding/tc-admin-api",
        "status_json": f"{ROOT}/AI_coding/logs/status_hqkdngay_daily.json",
        "status_jsonl": f"{ROOT}/AI_coding/logs/status_hqkdngay_daily.jsonl",
    },
    "prod": {
        "database_url": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard",
        "log_file": f"{ROOT}/AI_coding/logs/cron_hqkdngay_daily_prod.log",
        "disabled_flag": f"{ROOT}/AI_coding/logs/cron_hqkdngay_daily_prod.disabled",
        "verify_api_dir": "/home/itadmin/apps/tc-console/tc-admin-api",
        "status_json": f"{ROOT}/AI_coding/logs/status_hqkdngay_daily_prod.json",
        "status_jsonl": f"{ROOT}/AI_coding/logs/status_hqkdngay_daily_prod.jsonl",
    },
}
DEFAULT_ENV = "test"
LOG_FILE = ENVIRONMENTS[DEFAULT_ENV]["log_file"]
DISABLED_FLAG = ENVIRONMENTS[DEFAULT_ENV]["disabled_flag"]
DATABASE_URL = ENVIRONMENTS[DEFAULT_ENV]["database_url"]
VERIFY_API_DIR = ENVIRONMENTS[DEFAULT_ENV]["verify_api_dir"]

SCHEDULE_VN = "05:00 · 16:45 · 17:15 (prod sớm hơn 10')"   # 3 lượt/ngày (đổi 29/08/2026). Ghi vào artifact
# cho agent giám sát khỏi hard-code mốc giờ — KHÔNG ai parse chuỗi này, chỉ hiển thị.
# Lượt 05:00 sáng gánh các nguồn nộp sau giờ chiều (KSCL của XDV nộp 18:02-18:31 VN).
# Lượt prod chạy sau test 5 phút; xem khối chú thích trong crontab.

# 12 đơn vị PHẢI có báo cáo ngày — mẫu số của artifact. Lấy từ `_UNITS` của deriver để không có
# 2 nguồn sự thật; thêm/bớt đơn vị chỉ sửa 1 chỗ. Fallback hard-code cho trường hợp import vỡ
# (deriver kéo theo openpyxl/be_bridge): thà dùng danh sách cũ còn hơn mất mẫu số và im lặng.
_UNITS_FALLBACK = ("SRVF", "XANHVINHPHUC", "HTXXANHTUYENQUANG", "HTXXANHVINHPHUC", "ANTAXI",
                   "ANKHACHSAN", "GLOBALAI", "TRAMSAC", "DUAN", "HO", "HUNGTHINH", "XDV")

# Tên tiếng Việt cho tin gửi Ban lãnh đạo / KSNB — họ không đọc mã thư mục. Đặt ở đây (nguồn dữ
# liệu) thay vì trong chỉ dẫn của agent: agent lấy `ten` nguyên văn, không phải tự dịch.
UNIT_NAMES = {
    "SRVF": "Showroom Vinfast", "XDV": "Xưởng dịch vụ Vinfast", "TRAMSAC": "Trạm sạc Vgreen",
    "DUAN": "Dự án", "HO": "Khối hỗ trợ tập đoàn", "HUNGTHINH": "Xe tải Hưng Thịnh",
    "GLOBALAI": "Global AI", "ANTAXI": "An Taxi", "ANKHACHSAN": "An Khách sạn",
    "XANHVINHPHUC": "Xanh Vĩnh Phúc", "HTXXANHTUYENQUANG": "HTX Xanh Tuyên Quang",
    "HTXXANHVINHPHUC": "HTX Xanh Vĩnh Phúc",
}


def units_source() -> dict:
    """`_UNITS` của deriver — nguồn DUY NHẤT cho danh sách đơn vị kỳ vọng."""
    try:
        sys.path.insert(0, f"{AGENT}/scripts")
        import derive_hqkd_ngay                      # noqa: PLC0415
        return dict(derive_hqkd_ngay._UNITS)
    except Exception as ex:                          # noqa: BLE001
        log(f"  CANH BAO: khong import duoc _UNITS ({ex}) -> dung danh sach fallback")
        return {u: {} for u in _UNITS_FALLBACK}


def unit_names(src: dict) -> dict:
    """Tên tiếng Việt cho tin gửi lãnh đạo, có ĐƯỜNG LÙI.

    Thêm đơn vị mới vào `_UNITS` mà quên khai ở `UNIT_NAMES` thì trước đây tin hiện trơ mã thư mục
    ("VANTAIABC") giữa danh sách tên tiếng Việt. Giờ lùi về `khoi` của chính deriver (vd "Khối KD
    Vinfast - Showroom") — dài hơn tên rút gọn nhưng vẫn đọc được, và không ai phải nhớ sửa 2 chỗ
    mới thêm được một đơn vị.
    """
    return {u: UNIT_NAMES.get(u) or (src.get(u) or {}).get("khoi") or u for u in src}


REPORT_TYPES = ("baocaohqkdngay",)
# 4 report_type derive_hqkd_ngay.py ghi ra (RT_HQKD/RT_PNLT/RT_CHIPHI/RT_DTHU) — verify() đếm
# ngày trên bất kỳ loại nào có mặt (mỗi đơn vị/layout không chắc sinh đủ cả 4).
DAY_REPORT_TYPES = ("HQKD_D", "PNLT_D", "CHIPHI_D", "DTHU_D")
META_REFRESH_TIMEOUT = 240
META_REFRESH_POLL = 5
ARRIVE_TIMEOUT = 300
ARRIVE_POLL = 5


def log(msg: str):
    line = f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}"
    print(line, flush=True)
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def http(method: str, path: str, payload=None, timeout=20):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(RECEIVER_URL + path, data=data, method=method,
                                headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        body = r.read().decode(errors="replace")
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return {"raw": body[:200]}


def meta_mtime():
    try:
        return os.path.getmtime(AVAILABLE_META)
    except OSError:
        return None


def wait_metadata(before) -> bool:
    t0 = time.time()
    while time.time() - t0 < META_REFRESH_TIMEOUT:
        time.sleep(META_REFRESH_POLL)
        if meta_mtime() != before:
            time.sleep(1)
            log(f"  danh sách mới về sau {int(time.time() - t0)}s")
            return True
    log(f"  CANH BAO: {META_REFRESH_TIMEOUT}s máy Local chưa nộp danh sách mới (agent còn sống"
        " không?) -> dùng danh sách CŨ, createdAt dưới đây là của lần quét trước")
    return False


def read_metadata():
    for i in range(3):
        try:
            with open(AVAILABLE_META, encoding="utf-8") as fh:
                return json.load(fh)
        except json.JSONDecodeError:
            time.sleep(2)
        except OSError as ex:
            log(f"DỪNG: không đọc được {AVAILABLE_META} ({ex})")
            return None
    log(f"DỪNG: {AVAILABLE_META} đọc 3 lần vẫn lỗi JSON (ghi dở hoặc hỏng)")
    return None


def target_periods():
    """[(period, month, year)] = tháng NÀY + tháng TRƯỚC theo giờ VN (đầu tháng còn chốt sổ cũ)."""
    now = datetime.now(VN)
    prev = (now.replace(day=1) - timedelta(days=1))
    return [(f"{d.year}-{d.month:02d}", d.month, d.year) for d in (now, prev)]


def pick_targets(meta: list, periods: list) -> list:
    """Chọn entry metadata thuộc các kỳ cần kéo, MỌI đơn vị (không lọc company — 12 đơn vị đều
    nằm chung report_type 'baocaohqkdngay', xem docstring module)."""
    out = []
    for period, month, year in periods:
        for e in meta:
            if e.get("report_type") not in REPORT_TYPES or not e.get("fileName"):
                continue
            thang, nam_ten = cron_status.ky_cua_entry(e)
            if thang != month:
                continue
            years = {int(t) for t in re.findall(r"(20\d{2})", e["fileName"])}
            if years and year not in years:
                continue
            if nam_ten is not None and nam_ten != year:
                continue
            out.append({**e, "_period": period})
    return out



def sidecar_saved_at(entry: dict):
    p = os.path.join(RECEIVED_DIR, entry.get("company") or "", entry.get("report_type") or "",
                     os.path.splitext(entry["fileName"])[0] + ".json")
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh).get("saved_at")
    except (OSError, json.JSONDecodeError):
        return None


def xlsx_path(entry: dict) -> str:
    return os.path.join(RECEIVED_DIR, entry.get("company") or "",
                        entry.get("report_type") or "", entry["fileName"])


def source_id(entry: dict) -> str:
    """'<công_ty_thư_mục>::<tên_file>' — khớp source_file lưu trong raw_rows (xem
    derive_hqkd_ngay._source_id / Dashboard_Agent source_catalog.source_id_from_path)."""
    return f"{entry.get('company') or ''}::{entry['fileName']}"


def request_files(targets: list) -> dict:
    before = {}
    for e in targets:
        before[e["fileName"]] = sidecar_saved_at(e)
        payload = {k: e.get(k) for k in ("company", "report_type", "fileName", "path",
                                         "month", "periodType", "status")}
        try:
            http("POST", "/request-file", payload)
            log(f"  đã xin: [{e.get('company')}] {e['fileName'][:55]}")
        except (urllib.error.URLError, OSError) as ex:
            log(f"  LỖI xin file {e['fileName'][:50]}: {ex}")
    return before


def wait_arrival(targets: list, before: dict) -> list:
    deadline = time.time() + ARRIVE_TIMEOUT
    pending = {e["fileName"]: e for e in targets}
    arrived = []
    while pending and time.time() < deadline:
        time.sleep(ARRIVE_POLL)
        for fn, e in list(pending.items()):
            now_at = sidecar_saved_at(e)
            if now_at and now_at != before.get(fn):
                arrived.append(e)
                del pending[fn]
                log(f"  đã về: [{e.get('company')}] {fn[:55]} (saved_at {now_at})")
    for fn, e in pending.items():
        old = before.get(fn)
        log(f"  KHÔNG VỀ sau {ARRIVE_TIMEOUT}s: [{e.get('company')}] {fn[:50]}"
            + (f" — vẫn dùng bản cũ {old}" if old else " — chưa từng có bản nào"))
    return arrived


def _ky_moi_tao(js: dict) -> list:
    """Kỳ vừa được KHAI SINH trong lượt nạp này (servers/common/dataset_ky.py).

    Không phải cảnh báo — là SỰ KIỆN đáng ghi: kỳ mới xuất hiện trong ô chọn kỳ của dashboard,
    và nếu về sau có tranh cãi "kỳ này ở đâu ra" thì log là chỗ duy nhất trả lời được.
    """
    ks = set()
    for x in [js] + list(js.get("derived") or []) + list(js.get("processed") or []):
        v = x.get("ky_moi_tao") if isinstance(x, dict) else None
        ks.update([v] if isinstance(v, str) else (v or []))
    return sorted(ks)


def autofill(entry: dict):
    """agent_cli.py autofill: is_daily_report() gate tự dispatch sang derive_hqkd_ngay.derive()
    (DELETE-then-insert idempotent theo source_file, đọc lại TOÀN BỘ ngày có trong file — xem
    docstring module). Không dùng LLM.

    Trả (ok, ly_do): `ly_do` là câu error NGẮN lấy từ JSON dòng cuối (vd "đọc được 6 sheet ngày
    (D1..D6) nhưng không sheet nào ra chỉ tiêu") để artifact có lý do thật, không phải suy đoán.
    rc=0 nhưng JSON `ok:false` VẪN là lỗi nạp — agent_cli không đổi exit code cho ca này.
    """
    path = xlsx_path(entry)
    env = {**os.environ, "DATABASE_URL": DATABASE_URL}
    cmd = [AGENT_PY, "scripts/agent_cli.py", "autofill", path, "--period", entry["_period"]]
    p = subprocess.run(cmd, cwd=AGENT, env=env, capture_output=True, text=True, timeout=900)
    tail = (p.stdout or p.stderr or "").strip().splitlines()
    last = tail[-1] if tail else ""
    log(f"  autofill [{entry.get('company')}] {entry['_period']}: rc={p.returncode}"
        f" | {last[:280]}")
    if p.returncode != 0:
        return False, f"autofill lỗi (rc={p.returncode})"
    try:
        js = json.loads(last)
    except (json.JSONDecodeError, ValueError):
        return True, None            # không parse được thì để verify quyết định
    for _k in _ky_moi_tao(js):
        log(f"  KỲ MỚI: khai sinh kỳ {_k} (chưa từng có dataset; nguồn số thực tế)")
    if js.get("ok"):
        return True, None
    errs = [d.get("error") for d in (js.get("derived") or []) if d.get("error")]
    return False, (errs[0][:200] if errs else "autofill trả ok=false, không kèm lý do")


def verify(entry: dict, want_day: str):
    """Kiểm theo ĐÚNG source_file (mỗi đơn vị 1 file riêng, không gộp toàn kỳ như THUCHI):
    1. có ít nhất 1 report_type ngày (HQKD_D/PNLT_D/CHIPHI_D/DTHU_D) với dòng thật.
    2. max(ngay) >= want_day -> đã có số liệu ngày hôm qua (không phụ thuộc múi giờ nguồn).
    3. VÂN TAY của kỳ (số dòng + tổng |amount| làm tròn 6 chữ số) -> để lượt sau biết file có ĐỔI.

    Vì sao cần (3): với file dựng sẵn cột cả tháng, (1) và (2) vô nghĩa — XDV có 98 dòng mỗi ngày
    12→15/08 với tổng GIỐNG HỆT 1.655, HTX Xanh VP có số tới tận 20/08. "Có dòng cho ngày hôm qua"
    và cả "số khác 0" đều không chứng minh kế toán đã nhập. Bằng chứng duy nhất đáng tin là NỘI DUNG
    FILE ĐỔI so với lượt trước — họ update vào chính file tháng đó, giống bên dòng tiền."""
    # ĐƠN VỊ ĐÃ CHUYỂN NGUỒN: `derive_hqkd_ngay` bỏ mọi ngày >= `bo_tu_ngay` của đơn vị, nên với
    # kỳ nằm trọn sau mốc thì file tay ĐÚNG RA phải không còn dòng nào. Hỏi DB rồi kết luận
    # "không có dòng nào" là dựng một ô đỏ vĩnh viễn cho một việc bình thường. Chỉ ngắn mạch khi
    # mốc rơi đúng ngày 01 và kỳ nằm từ tháng đó trở đi — cutover giữa tháng thì tháng đó vẫn còn
    # phần trước mốc, phải kiểm như thường.
    moc = (units_source().get(entry.get("company")) or {}).get("bo_tu_ngay")
    if moc and moc.endswith("-01") and entry["_period"] >= moc[:7]:
        out = f"DA_CHUYEN_NGUON(tu {moc})"
        log(f"  kiem [{entry.get('company')}] {entry['_period']}: {out}"
            " — nguon tay da duoc thay bang nguon tu dong, khong con dong nao la DUNG")
        return parse_verify(out)
    sid = source_id(entry)
    rts = ",".join(f"'{t}'" for t in DAY_REPORT_TYPES)
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.database.session import get_db;"
        f"sid={sid!r};w={want_day!r};"
        f"r=get_db().execute(\"SELECT COUNT(DISTINCT ngay) nd,MAX(ngay) mx,COUNT(*) sd,"
        f"SUM(ABS(COALESCE(amount,0))) tg FROM raw_rows "
        f"WHERE source_file=? AND report_type IN ({rts})\",(sid,)).fetchone();"
        "print('KHONG_CO_DONG_NAO') if not r or not r['nd'] else "
        "print('so_ngay=%s max_ngay=%s van_tay=%s:%.6f %s'%(r['nd'],r['mx'],r['sd'],"
        "float(r['tg'] or 0),"
        "('OK_CO_NGAY_HOM_QUA' if (r['mx'] or '')>=w else 'THIEU_NGAY_HOM_QUA(can>=%s)'%w)))"
    )
    api = VERIFY_API_DIR
    p = subprocess.run([f"{api}/.venv/bin/python", "-c", code], cwd=api,
                       env={**os.environ, "DATABASE_URL": DATABASE_URL},
                       capture_output=True, text=True, timeout=180)
    out = (p.stdout or p.stderr).strip()
    log(f"  kiem [{entry.get('company')}] {entry['_period']}: {out[:200]}")
    if "THIEU_NGAY_HOM_QUA" in out:
        log(f"  CANH BAO [{entry.get('company')}]: chua co so lieu ngay hom qua -> co the file"
            " nguon chua duoc ke toan cap nhat, hoac cron chay som hon luc cap nhat.")
    return parse_verify(out)


def parse_verify(out: str) -> dict:
    """Bóc 'so_ngay=10 max_ngay=2026-08-10 OK_CO_NGAY_HOM_QUA' thành dict cho artifact.

    Bóc ngay tại nguồn thay vì để agent parse lại: đây là chuỗi do CHÍNH file này sinh ra, giữ
    việc đọc nó ở cùng chỗ thì đổi format không làm hỏng bên tiêu thụ.
    """
    d = {"raw": out[:200], "so_ngay": None, "max_ngay": None, "code": None}
    for code in ("DA_CHUYEN_NGUON", "KHONG_CO_DONG_NAO", "THIEU_NGAY_HOM_QUA",
                 "OK_CO_NGAY_HOM_QUA"):
        if code in out:
            d["code"] = code
            break
    m = re.search(r"so_ngay=(\d+)", out)
    if m:
        d["so_ngay"] = int(m.group(1))
    m = re.search(r"max_ngay=(\d{4}-\d{2}-\d{2})", out)
    if m:
        d["max_ngay"] = m.group(1)
    m = re.search(r"van_tay=(\d+:[\d.]+)", out)
    if m:
        d["van_tay"] = m.group(1)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="chỉ in file sẽ kéo, KHÔNG xin file / KHÔNG nạp DB")
    ap.add_argument("--env", choices=list(ENVIRONMENTS), default=DEFAULT_ENV,
                    help=f"đích ghi dữ liệu (mặc định '{DEFAULT_ENV}')")
    args = ap.parse_args()

    global LOG_FILE, DISABLED_FLAG, DATABASE_URL, VERIFY_API_DIR
    cfg = ENVIRONMENTS[args.env]
    LOG_FILE, DISABLED_FLAG, DATABASE_URL, VERIFY_API_DIR = (
        cfg["log_file"], cfg["disabled_flag"], cfg["database_url"], cfg["verify_api_dir"])
    # DRY-RUN ghi log SANG FILE RIÊNG: log chính là nguồn duy nhất cho panel "Chạy tự động theo lịch"
    # (source_bridge._job_last_run đọc khối cuối; không thấy dòng XONG là kết luận "chạy dở / lỗi
    # không rõ"). Một lượt dry-run vào log chính làm job hiện ĐỎ OAN cho tới lần chạy thật kế tiếp —
    # gặp thật 2026-08-12, panel prod đỏ cả 2 job trong khi dữ liệu hoàn toàn bình thường.
    if args.dry_run:
        LOG_FILE = LOG_FILE[:-4] + "_dryrun.log" if LOG_FILE.endswith(".log") else LOG_FILE

    log("=" * 70)
    log(f"MOI TRUONG: {args.env} (DB {DATABASE_URL.rsplit('@', 1)[-1]})")
    # Artifact JSON cho agent giám sát (xem cron_status.py). DRY-RUN không ghi: nó không kéo/không
    # nạp nên trạng thái sinh ra là giả, ghi vào sẽ đè mất trạng thái thật của lượt trước.
    _src = units_source()
    st = None if args.dry_run else cron_status.StatusWriter(
        job="hqkdngay_daily", env=args.env, json_path=cfg["status_json"],
        jsonl_path=cfg["status_jsonl"], schedule_vn=SCHEDULE_VN, expected=sorted(_src),
        names=unit_names(_src))

    def done(status=cron_status.RUN_OK, note=None, rc=0):
        """Mọi đường ra của main() phải đi qua đây: artifact THIẾU bị agent hiểu là 'cron không
        chạy', nên thoát im lặng ở nhánh lỗi sẽ thành báo động sai."""
        if st:
            st.finish(status, note)
        return rc

    if os.path.exists(DISABLED_FLAG) and not args.dry_run:
        log("BO QUA: dang TAT tu giao dien (Nguon du lieu > Keo tu dong). Xoa co de bat lai.")
        return done(cron_status.RUN_TAT, "job dang bi TAT tu giao dien Nguon du lieu")
    periods = target_periods()
    log(f"BẮT ĐẦU — kỳ cần kéo: {[p for p, _, _ in periods]}"
        + (" [DRY-RUN]" if args.dry_run else ""))
    if st:
        st.set_run(periods=[p for p, _, _ in periods], period_chinh=periods[0][0])

    before_meta = meta_mtime()
    try:
        http("POST", "/request-metadata")
    except (urllib.error.URLError, OSError) as ex:
        log(f"DỪNG: không gọi được receiver {RECEIVER_URL} ({ex}). receiver_server còn chạy không?")
        return done(cron_status.RUN_DUNG_SOM, f"không gọi được receiver: {str(ex)[:120]}", 1)
    log(f"đã xin làm mới danh sách, chờ tối đa {META_REFRESH_TIMEOUT}s cho tới khi danh sách đổi")
    if args.dry_run:
        log("DRY-RUN: không chờ, đọc luôn danh sách đang có (có thể là bản quét cũ)")
    else:
        wait_metadata(before_meta)

    meta = read_metadata()
    if meta is None:
        return done(cron_status.RUN_DUNG_SOM, "không đọc được available_metadata.json", 1)

    targets = pick_targets(meta, periods)
    for e in meta:                       # nổ TO khi tên file mới lại không suy được kỳ
        if e.get("report_type") in REPORT_TYPES and e.get("fileName") \
                and cron_status.ky_cua_entry(e)[0] is None:
            log(f"  CANH BAO [{e.get('company')}]: co file o nguon nhung khong doc duoc KY tu ten"
                f" '{e['fileName'][:60]}' -> khong the xep vao ky nao de keo, can sua ten file"
                " hoac bo sung dang ten vao cron_status.thang_tu_ten_file()")
    if not targets:
        log("DỪNG: không có file nào khớp kỳ cần kéo (chưa đơn vị nào tạo file tháng này?)")
        return done(cron_status.RUN_DUNG_SOM, "không có file nào khớp kỳ cần kéo", 1)
    for e in targets:
        log(f"  chọn [{e['_period']}] [{e.get('company')}] {e['fileName'][:52]}"
            f" createdAt={e.get('createdAt')}")
        if st and e["_period"] == periods[0][0]:
            st.record(e.get("company") or "?", file=e["fileName"],
                      nguon_sua_luc=e.get("createdAt"), ky=e["_period"],
                      ly_do="đã chọn để kéo, chưa có kết quả nạp")
    if args.dry_run:
        log("DRY-RUN: dừng ở đây")
        return 0

    log(f"KÉO {len(targets)} báo cáo ngày ({len({e.get('company') for e in targets})} đơn vị)")
    before = request_files(targets)
    arrived_names = {e["fileName"] for e in wait_arrival(targets, before)}

    # Ngày số liệu đòi hỏi = hôm qua, mọi ngày như nhau (kể cả Chủ nhật). Xem
    # cron_status.ngay_can() để biết vì sao không có ngoại lệ cuối tuần.
    yday = cron_status.ngay_can(datetime.now(VN))
    today = datetime.now(VN).strftime("%Y-%m-%d")
    ky_chinh = periods[0][0]
    ok = 0
    # kỳ cũ trước, kỳ mới sau — không có phụ thuộc chéo giữa các ĐƠN VỊ nên thứ tự công ty tự do.
    for e in sorted(targets, key=lambda x: (x["_period"], x.get("company") or "")):
        # Artifact chỉ theo KỲ CHÍNH (tháng này): kỳ tháng trước đã chốt, đưa vào cùng mẫu số 12
        # đơn vị sẽ có 2 bản ghi cho 1 đơn vị và trạng thái của tháng cũ đè trạng thái tháng này.
        rec = st.record if (st and e["_period"] == ky_chinh) else (lambda *a, **k: None)
        unit = e.get("company") or "?"
        if not os.path.exists(xlsx_path(e)):
            log(f"  BỎ QUA [{e.get('company')}] {e['_period']}: không có file trên đĩa")
            rec(unit, state=cron_status.STATE_CHUA_CO_FILE, arrived=False,
                ly_do="đã xin file nhưng không về, trên đĩa cũng chưa có bản nào")
            continue
        arrived = e["fileName"] in arrived_names
        af_ok, af_err = autofill(e)
        if not af_ok:
            # `ly_do` là câu cho LÃNH ĐẠO (agent dùng nguyên văn), `ly_do_ky_thuat` giữ nguyên câu
            # error của deriver cho IT. Trộn hai thứ vào một trường thì tên sheet (D1..D6) lọt vào
            # tin gửi lãnh đạo, và agent buộc phải tự diễn giải -> nó thêm suy đoán không có trong
            # dữ liệu (đã gặp thật: "file còn để trống hoặc ghi 0 ở các mã tổng").
            rec(unit, state=cron_status.STATE_LOI_NAP, arrived=arrived,
                ly_do="hệ thống không đọc được số liệu trong file, cần kế toán kiểm tra lại file",
                ly_do_ky_thuat=af_err)
            continue
        ok += 1
        vr = verify(e, yday if yday.startswith(e["_period"]) else "0000-00-00")
        van_tay_cu = st.van_tay_cu.get(unit) if st else None
        state, doi_luc = cron_status.state_from_verify(
            vr, af_ok, today, van_tay_cu, st.doi_luc_cu.get(unit) if st else None)
        rec(unit, state=state, arrived=arrived, so_ngay=vr.get("so_ngay"),
            max_ngay=vr.get("max_ngay"), verify_code=vr.get("code"),
            van_tay=vr.get("van_tay"), doi_luc=doi_luc,
            ly_do=_ly_do(state, vr, yday, van_tay_cu, doi_luc, today))

    log(f"XONG — nạp thành công {ok}/{len(targets)} báo cáo ngày")
    if st:
        st.set_run(nap_thanh_cong=ok, so_file_keo=len(targets))
    return done(cron_status.RUN_OK, rc=0 if ok else 1)


def _ly_do(state: str, vr: dict, ngay_can: str, van_tay_cu: str = None,
           doi_luc: str = None, today: str = None) -> str:
    """Câu lý do bằng lời thường, viết SẴN ở đây cho agent dùng nguyên văn — agent không phải diễn
    giải mã kỹ thuật (`KHONG_CO_DONG_NAO`, `max_ngay=`) thành tiếng Việt và không thể diễn giải sai."""
    if state == cron_status.STATE_DU:
        d = vr.get("max_ngay") or ngay_can
        return f"đã có số liệu đến ngày {d[8:10]}/{d[5:7]}"
    # Câu lý do phải NGẮN: mẫu tin (user chốt 2026-08-12) dồn cả mục lên 1 dòng và gom các đơn vị
    # cùng lý do vào một cụm, nên lý do dài làm tràn dòng trên điện thoại. Rút ở đây thay vì để
    # agent tự cắt — agent cắt là agent diễn giải, và diễn giải thì có ngày sai.
    if state == cron_status.STATE_CHAM:
        mx = vr.get("max_ngay")
        # Ca RIÊNG: file có đủ dòng tới ngày cần nhưng vân tay y hệt lượt trước -> không ai nhập gì.
        # Nói "mới nhất <ngày>, chậm N ngày" ở đây là sai sự thật, vì ngày đó CÓ dòng — chỉ là dòng cũ.
        if van_tay_cu and doi_luc != today:
            return ("file chưa đổi kể từ lượt trước, kế toán chưa cập nhật"
                    + (f" (lần cập nhật gần nhất {doi_luc[8:10]}/{doi_luc[5:7]})" if doi_luc else ""))
        if mx:
            tre = (datetime.strptime(ngay_can, "%Y-%m-%d") - datetime.strptime(mx, "%Y-%m-%d")).days
            return f"mới nhất {mx[8:10]}/{mx[5:7]}, chậm {tre} ngày"
        return "chưa có số liệu của ngày cần"
    if state == cron_status.STATE_KHONG_XAC_NHAN:
        return "file dựng sẵn cả tháng, chưa xác nhận được"
    if state == cron_status.STATE_DA_CHUYEN_NGUON:
        return "đã chuyển sang nguồn tự động, file tay không còn là nguồn của kỳ này"
    if state == cron_status.STATE_LOI_NAP:
        return "hệ thống không đọc được số liệu trong file"
    return "chưa thấy file báo cáo ở nguồn"


if __name__ == "__main__":
    sys.exit(main())

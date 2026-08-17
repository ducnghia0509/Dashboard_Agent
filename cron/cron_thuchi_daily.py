#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KÉO + NẠP DÒNG TIỀN THEO NGÀY tự động — chạy 13:00 giờ VN mỗi ngày (cron 06:00 UTC).

Kế toán cập nhật file trên VPS lúc 11:00 giờ VN của ngày n+1 (thêm số liệu ngày n). Script này
chạy 13:00 để kéo bản mới về rồi nạp lại dashboard TEST.

MỐC TRONG CRONTAB PHẢI VIẾT THEO UTC (`0 6 * * *`). Cron của Ubuntu (3.0pl1) bỏ qua CRON_TZ khi
tính lịch — man 5 crontab / LIMITATIONS: "does not support per-user timezones". Đặt
CRON_TZ=Asia/Ho_Chi_Minh + `0 13` thì job nổ 13:00 UTC = 20:00 giờ VN, trễ 7 tiếng.

Luồng:
  1. POST /request-metadata           -> bảo máy Local quét & gửi danh sách; CHỜ tới khi
     available_metadata.json thực sự đổi (mtime), không sleep cứng
  2. POST /request-file  (EP FORCE)   -> tháng NÀY + tháng TRƯỚC, baocaothuchi (+ baocaonganhang
     nếu PULL_NGANHANG bật — hiện TẮT, xem khối ghi chú ở PULL_NGANHANG)
  3. chờ file thực sự về (sidecar saved_at đổi), có timeout
  4. agent_cli.py autofill <file thu chi>  -> extract_tien (THUCHI/SDT/VAY theo NGÀY) + chain VAY
  5. đối chiếu Σ các ngày == tổng kỳ, ghi log

VÌ SAO KHÔNG DÙNG `sync_orchestrator.py pull`: cmd_plan chỉ xin file CHƯA nhận, khoá theo
(company, fileName). File tháng giữ NGUYÊN TÊN suốt tháng nên sau lần kéo đầu nó bị coi "đã nhận"
vĩnh viễn -> bản cập nhật hàng ngày không bao giờ về. Ở đây POST /request-file trực tiếp từ
available_metadata.json để ép kéo lại. receiver_server ghi file bằng "wb" nên ghi đè an toàn.

BÁO CÁO NGÂN HÀNG phải kéo TRƯỚC file thu chi: chain VAY (extract_no_den_han / extract_lai_vay /
extract_vay_kyhan, chạy bên trong autofill) glob thẳng THUCHI/baocaonganhang/*.xlsx từ ĐĨA.

Chạy tay: Dashboard_Agent/.venv/bin/python cron/cron_thuchi_daily.py [--dry-run] [--with-nganhang]

VỊ TRÍ: xem khối cùng tên trong `cron_hqkdngay_daily.py` — script ở repo `Dashboard_Agent`, log và
artifact vẫn ở `AI_coding/logs/` (cố ý, vì agent openclaw và `source_bridge._LOGS` trỏ vào đó).
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
# Theo --env (mặc định 'test', GIỮ NGUYÊN hành vi cũ khi chạy không kèm cờ — crontab hiện tại của
# TEST gọi script KHÔNG có --env). 'prod' ghi log/cờ tắt RIÊNG (không lẫn với test) + venv verify
# lấy từ đúng checkout prod (apps/tc-console/tc-admin-api) để khỏi lệch code kiểm tra.
ENVIRONMENTS = {
    "test": {
        # DB TEST (5435). Mật khẩu có '$' -> phải %24 trong URL.
        "database_url": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
        "log_file": f"{ROOT}/AI_coding/logs/cron_thuchi_daily.log",
        "disabled_flag": f"{ROOT}/AI_coding/logs/cron_thuchi_daily.disabled",
        "verify_api_dir": f"{ROOT}/AI_coding/tc-admin-api",
        "status_json": f"{ROOT}/AI_coding/logs/status_thuchi_daily.json",
        "status_jsonl": f"{ROOT}/AI_coding/logs/status_thuchi_daily.jsonl",
    },
    "prod": {
        "database_url": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard",
        "log_file": f"{ROOT}/AI_coding/logs/cron_thuchi_daily_prod.log",
        "disabled_flag": f"{ROOT}/AI_coding/logs/cron_thuchi_daily_prod.disabled",
        "verify_api_dir": "/home/itadmin/apps/tc-console/tc-admin-api",
        "status_json": f"{ROOT}/AI_coding/logs/status_thuchi_daily_prod.json",
        "status_jsonl": f"{ROOT}/AI_coding/logs/status_thuchi_daily_prod.jsonl",
    },
}
DEFAULT_ENV = "test"
# Cờ TẮT do UI đặt (Nguồn dữ liệu > Kéo tự động). Dùng file thay vì sửa crontab từ web:
# tiến trình API không phải chỉnh crontab của user, và lịch cron giữ nguyên để bật lại tức thì.
LOG_FILE = ENVIRONMENTS[DEFAULT_ENV]["log_file"]
DISABLED_FLAG = ENVIRONMENTS[DEFAULT_ENV]["disabled_flag"]
DATABASE_URL = ENVIRONMENTS[DEFAULT_ENV]["database_url"]
VERIFY_API_DIR = ENVIRONMENTS[DEFAULT_ENV]["verify_api_dir"]

SCHEDULE_VN = "13:00"       # ghi vào artifact để agent giám sát không phải hard-code mốc giờ
REPORT_TYPES = ("baocaothuchi", "baocaonganhang")
# TẠM TẮT KÉO BÁO CÁO NGÂN HÀNG (user chốt 2026-08-11): hiện chỉ cần báo cáo thu chi, mà 11/08 cả
# 10 file ngân hàng đều KHÔNG VỀ sau 300s -> mỗi lượt chạy tốn thêm 5 phút chờ vô ích và log đầy
# cảnh báo giả. Bật lại: đổi True (hoặc chạy 1 lượt với --with-nganhang).
#
# TẮT CÁI GÌ VÀ KHÔNG TẮT CÁI GÌ — chỉ tắt VIỆC KÉO, không tắt chain VAY. extract_no_den_han /
# extract_lai_vay / extract_vay_kyhan vẫn chạy trong autofill và vẫn glob file ngân hàng TỪ ĐĨA
# (bản kéo về gần nhất). Nên số vay KHÔNG mất, chỉ ĐỨNG YÊN ở bản đĩa hiện có:
#   - nợ gốc đến hạn  -> payload.den_han  -> thẻ denHan/tlDenHan/bankDenHan (màn Dòng tiền/Vay)
#   - lãi vay trong kỳ -> amount2         -> thẻ "Lãi vay"
#   - tách kỳ hạn Ngắn/Trung hạn -> dim2
# Cả 3 đều theo THÁNG (khoá period_month, cong_ty, dim1=ngân hàng), KHÔNG theo ngày — đó là lý do
# kéo hằng ngày chỉ có tác dụng bắt kịp lúc kế toán sửa file GIỮA tháng, chứ không sinh số liệu
# ngày mới nào. TUYỆT ĐỐI KHÔNG xoá file ngân hàng trên đĩa khi đang tắt: chain sẽ không thấy
# nguồn và các thẻ trên rỗng đi.
PULL_NGANHANG = False
META_REFRESH_TIMEOUT = 240  # giây, chờ danh sách MỚI từ máy Local (đo thực tế: quét 276 file ~70s)
META_REFRESH_POLL = 5
ARRIVE_TIMEOUT = 300        # giây, chờ tất cả file được yêu cầu về đĩa
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
    """Chờ máy Local nộp DANH SÁCH MỚI — nhận biết bằng mtime của available_metadata.json đổi.

    KHÔNG sleep cứng 20s như trước: đo thực tế agent mất ~70s để quét 276 file rồi upload, nên
    script đọc trúng bản quét TRƯỚC. Nội dung file tải về không bị ảnh hưởng (agent đọc file từ
    đĩa lúc upload), nhưng `createdAt` in ra log là của lần quét cũ -> chẩn đoán "kế toán chưa
    cập nhật" bị sai lệch. Log là thứ người ta tin khi đi tìm nguyên nhân, nó phải nói đúng."""
    t0 = time.time()
    while time.time() - t0 < META_REFRESH_TIMEOUT:
        time.sleep(META_REFRESH_POLL)
        if meta_mtime() != before:
            time.sleep(1)   # receiver mở file bằng "w" rồi mới json.dump -> chờ nó ghi xong
            log(f"  danh sách mới về sau {int(time.time() - t0)}s")
            return True
    log(f"  CANH BAO: {META_REFRESH_TIMEOUT}s máy Local chưa nộp danh sách mới (agent còn sống"
        " không?) -> dùng danh sách CŨ, createdAt dưới đây là của lần quét trước")
    return False


def read_metadata():
    """Đọc available_metadata.json, thử lại nếu trúng lúc receiver đang ghi dở. None = chịu."""
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
    """[(period, month, year)] = tháng NÀY + tháng TRƯỚC theo giờ VN.

    Kéo cả tháng trước vì đầu tháng kế toán còn đang chốt sổ tháng cũ."""
    now = datetime.now(VN)
    prev = (now.replace(day=1) - timedelta(days=1))
    return [(f"{d.year}-{d.month:02d}", d.month, d.year) for d in (now, prev)]


def pick_targets(meta: list, periods: list) -> list:
    """Chọn entry metadata thuộc các kỳ cần kéo.

    Khoá theo field `month` của metadata (đồng nhất cho mọi loại báo cáo — tên file thì mỗi bên
    một kiểu: 'THÁNG 08 NĂM 2026', 'T7.2026', 'M.202607'). `month` KHÔNG mang năm nên thêm chặn:
    nếu tên file có số 4-chữ-số trông như năm thì phải khớp năm mong đợi (tránh vắt sang năm khác
    ở mốc tháng 1 -> tháng 12)."""
    out = []
    for period, month, year in periods:
        for e in meta:
            if e.get("report_type") not in REPORT_TYPES or not e.get("fileName"):
                continue
            if e.get("month") != month:
                continue
            years = {int(t) for t in re.findall(r"(20\d{2})", e["fileName"])}
            if years and year not in years:
                continue
            out.append({**e, "_period": period})
    return out


def sidecar_saved_at(entry: dict):
    """saved_at trong sidecar .json cạnh file đã nhận; None nếu chưa từng nhận."""
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


def request_files(targets: list) -> dict:
    """POST /request-file cho từng target. Trả {fileName: saved_at TRƯỚC khi kéo} để bước sau
    biết file nào đã thực sự được ghi lại."""
    before = {}
    for e in targets:
        before[e["fileName"]] = sidecar_saved_at(e)
        payload = {k: e.get(k) for k in ("company", "report_type", "fileName", "path",
                                         "month", "periodType", "status")}
        try:
            http("POST", "/request-file", payload)
            log(f"  đã xin: {e['fileName'][:60]}")
        except (urllib.error.URLError, OSError) as ex:
            log(f"  LỖI xin file {e['fileName'][:50]}: {ex}")
    return before


def wait_arrival(targets: list, before: dict) -> list:
    """Chờ tới khi saved_at đổi (hoặc file mới xuất hiện). Trả danh sách target ĐÃ về.

    Bắt buộc chờ: /request-file chỉ XẾP HÀNG, máy Local poll 2s rồi mới upload. Không chờ thì
    bước autofill sẽ xử lý bản CŨ mà vẫn báo thành công."""
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
                log(f"  đã về: {fn[:60]} (saved_at {now_at})")
    for fn, e in pending.items():
        old = before.get(fn)
        log(f"  KHÔNG VỀ sau {ARRIVE_TIMEOUT}s: {fn[:55]}"
            + (f" — vẫn dùng bản cũ {old}" if old else " — chưa từng có bản nào"))
    return arrived


def autofill(entry: dict):
    """agent_cli.py autofill: extract_tien (THUCHI/SDT/VAY theo ngày) + chain VAY. Không dùng LLM.

    Trả (ok, ly_do) — xem docstring cùng tên trong cron_hqkdngay_daily.py: rc=0 kèm JSON `ok:false`
    vẫn là lỗi nạp, và `ly_do` lấy nguyên câu error để artifact có lý do thật.
    """
    path = xlsx_path(entry)
    env = {**os.environ, "DATABASE_URL": DATABASE_URL, "THUCHI_DAILY": "1"}
    cmd = [AGENT_PY, "scripts/agent_cli.py", "autofill", path, "--period", entry["_period"]]
    p = subprocess.run(cmd, cwd=AGENT, env=env, capture_output=True, text=True, timeout=900)
    tail = (p.stdout or p.stderr or "").strip().splitlines()
    last = tail[-1] if tail else ""
    log(f"  autofill {entry['_period']}: rc={p.returncode} | {last[:300]}")
    if p.returncode != 0:
        return False, f"autofill lỗi (rc={p.returncode})"
    try:
        js = json.loads(last)
    except (json.JSONDecodeError, ValueError):
        return True, None
    if js.get("ok"):
        return True, None
    return False, (str(js.get("error"))[:200] if js.get("error")
                   else "autofill trả ok=false, không kèm lý do")


def verify(period: str, want_day: str):
    """Kiểm 2 điều kiện thành công, ghi vào log:

    1. Σ các NGÀY == tổng kỳ  -> nạp theo ngày không làm lệch số tháng.
    2. max(ngay) >= want_day  -> ĐÃ có số liệu của ngày hôm qua. Đây là phép kiểm quan trọng nhất
       và KHÔNG phụ thuộc múi giờ: `createdAt` do máy Local gửi nên không rõ UTC hay giờ VN, không
       suy được "kế toán đã cập nhật chưa". Thiếu ngày hôm qua = chạy quá sớm -> phải dịch mốc cron.
    """
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.repositories import dataset_repository as dr;"
        "from app.database.session import get_db;"
        f"p={period!r};w={want_day!r};"
        "ids={d['period']:d['id'] for d in dr.list_all('month')};ds=ids.get(p);"
        "print('KHONG_CO_DATASET') if not ds else None;"
        "r=get_db().execute(\"SELECT COUNT(DISTINCT ngay) nd,MAX(ngay) mx,"
        "SUM(CASE WHEN dim1 LIKE ? THEN amount ELSE 0 END) thu,"
        "SUM(CASE WHEN dim1 LIKE ? THEN amount ELSE 0 END) chi"
        " FROM raw_rows WHERE report_type='THUCHI' AND dataset_id=?\",('A%','B%',ds)).fetchone()"
        " if ds else None;"
        "print('so_ngay=%s max_ngay=%s thu=%.6f chi=%.6f %s'%(r['nd'],r['mx'],"
        "float(r['thu'] or 0),float(r['chi'] or 0),"
        "('OK_CO_NGAY_HOM_QUA' if (r['mx'] or '')>=w else 'THIEU_NGAY_HOM_QUA(can>=%s)'%w)))"
        " if r else None"
    )
    api = VERIFY_API_DIR
    p = subprocess.run([f"{api}/.venv/bin/python", "-c", code], cwd=api,
                       env={**os.environ, "DATABASE_URL": DATABASE_URL},
                       capture_output=True, text=True, timeout=180)
    out = (p.stdout or p.stderr).strip()
    log(f"  kiem {period}: {out[:220]}")
    if "THIEU_NGAY_HOM_QUA" in out:
        log("  CANH BAO: chua co so lieu ngay hom qua -> co the cron chay som hon luc ke toan"
            " cap nhat. Xem createdAt o tren, can thi doi moc cron.")
    return parse_verify(out)


def parse_verify(out: str) -> dict:
    """Bóc chuỗi verify thành dict cho artifact — bản song sinh của hàm cùng tên trong
    cron_hqkdngay_daily.py, khác ở chỗ có thêm mã KHONG_CO_DATASET của riêng job thu chi."""
    d = {"raw": out[:220], "so_ngay": None, "max_ngay": None, "code": None}
    for code in ("KHONG_CO_DATASET", "THIEU_NGAY_HOM_QUA", "OK_CO_NGAY_HOM_QUA"):
        if code in out:
            d["code"] = code
            break
    m = re.search(r"so_ngay=(\d+)", out)
    if m:
        d["so_ngay"] = int(m.group(1))
    m = re.search(r"max_ngay=(\d{4}-\d{2}-\d{2})", out)
    if m:
        d["max_ngay"] = m.group(1)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="chỉ in file sẽ kéo, KHÔNG xin file / KHÔNG nạp DB")
    ap.add_argument("--env", choices=list(ENVIRONMENTS), default=DEFAULT_ENV,
                    help=f"đích ghi dữ liệu (mặc định '{DEFAULT_ENV}')")
    ap.add_argument("--with-nganhang", action="store_true",
                    help="kéo CẢ báo cáo ngân hàng lượt này (mặc định TẮT, xem PULL_NGANHANG)")
    args = ap.parse_args()
    pull_bank = PULL_NGANHANG or args.with_nganhang

    global LOG_FILE, DISABLED_FLAG, DATABASE_URL, VERIFY_API_DIR
    cfg = ENVIRONMENTS[args.env]
    LOG_FILE, DISABLED_FLAG, DATABASE_URL, VERIFY_API_DIR = (
        cfg["log_file"], cfg["disabled_flag"], cfg["database_url"], cfg["verify_api_dir"])
    # DRY-RUN ghi log sang file riêng — xem chú thích cùng chỗ trong cron_hqkdngay_daily.py: dry-run
    # ghi vào log chính sẽ làm panel "Chạy tự động theo lịch" báo "chạy dở / lỗi không rõ" oan.
    if args.dry_run:
        LOG_FILE = LOG_FILE[:-4] + "_dryrun.log" if LOG_FILE.endswith(".log") else LOG_FILE

    log("=" * 70)
    log(f"MOI TRUONG: {args.env} (DB {DATABASE_URL.rsplit('@', 1)[-1]})")
    periods = target_periods()
    ky_chinh = periods[0][0]
    # Mẫu số của job này là KỲ, không phải đơn vị: chỉ có 1 file "Báo cáo tiền tập đoàn" và kế toán
    # nhập số từng ngày thẳng vào file THÁNG. Kỳ tháng trước cũng được ghi nhưng `ky_chinh=False`
    # nên không tính vào summary (xem cron_status.finish).
    st = None if args.dry_run else cron_status.StatusWriter(
        job="thuchi_daily", env=args.env, json_path=cfg["status_json"],
        jsonl_path=cfg["status_jsonl"], schedule_vn=SCHEDULE_VN,
        expected=[p for p, _, _ in periods])

    def done(status=cron_status.RUN_OK, note=None, rc=0):
        """Mọi đường ra phải đi qua đây — xem docstring cùng tên ở cron_hqkdngay_daily.py."""
        if st:
            st.finish(status, note)
        return rc

    if st:
        st.set_run(periods=[p for p, _, _ in periods], period_chinh=ky_chinh,
                   keo_nganhang=pull_bank)
        for p, _, _ in periods[1:]:
            st.record(p, ky_chinh=False)
    if os.path.exists(DISABLED_FLAG) and not args.dry_run:
        log("BO QUA: dang TAT tu giao dien (Nguon du lieu > Keo tu dong). Xoa co de bat lai.")
        return done(cron_status.RUN_TAT, "job dang bi TAT tu giao dien Nguon du lieu")
    log(f"BẮT ĐẦU — kỳ cần kéo: {[p for p, _, _ in periods]}"
        + (" [DRY-RUN]" if args.dry_run else ""))

    before_meta = meta_mtime()      # LẤY TRƯỚC khi POST, không thì có thể bỏ lỡ lần ghi
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
    if not targets:
        log("DỪNG: không có file nào khớp kỳ cần kéo (kế toán chưa tạo file tháng này?)")
        return done(cron_status.RUN_DUNG_SOM, "không có file nào khớp kỳ cần kéo", 1)
    for e in targets:
        log(f"  chọn [{e['_period']}] {e['report_type']}/{e['fileName'][:52]}"
            f" createdAt={e.get('createdAt')}")
    if args.dry_run:
        log("DRY-RUN: dừng ở đây")
        return 0

    # Ngân hàng TRƯỚC thu chi: chain VAY đọc file ngân hàng từ đĩa khi autofill file thu chi.
    bank = [e for e in targets if e["report_type"] == "baocaonganhang"]
    cash = [e for e in targets if e["report_type"] == "baocaothuchi"]
    if not pull_bank:
        # Ghi RA LOG bản đĩa đang dùng cho từng file (không im lặng bỏ qua): chain VAY vẫn đọc
        # chúng, nên phải thấy được số vay đang "đứng" ở bản ngày nào.
        log(f"BỎ KÉO {len(bank)} báo cáo ngân hàng (PULL_NGANHANG=False) — chain VAY vẫn dùng bản"
            " trên đĩa, số vay/lãi/đến hạn giữ nguyên theo các bản dưới đây:")
        for e in sorted(bank, key=lambda x: x["fileName"]):
            on_disk = sidecar_saved_at(e)
            log(f"    [{e['_period']}] {e['fileName'][:52]} — "
                + (f"bản đĩa {on_disk}" if on_disk else "CHƯA CÓ BẢN NÀO trên đĩa"))
        bank = []

    log(f"KÉO {len(bank)} báo cáo ngân hàng + {len(cash)} báo cáo thu chi")
    before = request_files(bank + cash)
    arrived_names = {e["fileName"] for e in wait_arrival(bank + cash, before)}

    # Ngày hôm qua (giờ VN) — mục tiêu của lần chạy này: file phải đã có số liệu của ngày đó.
    # Ngày số liệu được phép đòi — KHÔNG phải cứ "hôm qua": lượt Chủ nhật lùi về thứ Sáu
    # vì không ai nhập ngày Chủ nhật. Xem cron_status.ngay_can().
    yday = cron_status.ngay_can(datetime.now(VN))
    today = datetime.now(VN).strftime("%Y-%m-%d")
    ok = 0
    for e in sorted(cash, key=lambda x: x["_period"]):     # kỳ cũ trước, kỳ mới sau
        ky = e["_period"]
        rec = st.record if st else (lambda *a, **k: None)
        la_ky_chinh = ky == ky_chinh
        if not os.path.exists(xlsx_path(e)):
            log(f"  BỎ QUA {ky}: không có file trên đĩa")
            rec(ky, state=cron_status.STATE_CHUA_CO_FILE, ky_chinh=la_ky_chinh, arrived=False,
                ten=f"Báo cáo tiền tập đoàn tháng {ky[5:7]}",
                ly_do="đã xin file nhưng không về, trên đĩa cũng chưa có bản nào")
            continue
        arrived = e["fileName"] in arrived_names
        af_ok, af_err = autofill(e)
        if not af_ok:
            # Tách lý do lãnh đạo / lý do kỹ thuật — xem chú thích cùng chỗ trong cron_hqkdngay_daily.py.
            rec(ky, state=cron_status.STATE_LOI_NAP, ky_chinh=la_ky_chinh, arrived=arrived,
                ten=f"Báo cáo tiền tập đoàn tháng {ky[5:7]}",
                ly_do="hệ thống không đọc được số liệu trong file, cần kế toán kiểm tra lại file",
                ly_do_ky_thuat=af_err)
            continue
        ok += 1
        # chỉ đòi "có ngày hôm qua" ở kỳ CHỨA ngày đó; kỳ trước thì bỏ qua điều kiện này
        vr = verify(ky, yday if yday.startswith(ky) else "0000-00-00")
        state = cron_status.state_from_verify(vr, af_ok, today)
        rec(ky, state=state, ky_chinh=la_ky_chinh, arrived=arrived, file=e["fileName"],
            nguon_sua_luc=e.get("createdAt"), so_ngay=vr.get("so_ngay"),
            max_ngay=vr.get("max_ngay"), verify_code=vr.get("code"),
            ten=f"Báo cáo tiền tập đoàn tháng {ky[5:7]}",
            ly_do=_ly_do(state, vr, yday, la_ky_chinh))

    log(f"XONG — nạp thành công {ok}/{len(cash)} kỳ thu chi")
    if st:
        st.set_run(nap_thanh_cong=ok, so_ky_keo=len(cash))
    return done(cron_status.RUN_OK, rc=0 if ok else 1)


def _ly_do(state: str, vr: dict, ngay_can: str, la_ky_chinh: bool) -> str:
    """Câu lý do bằng lời thường cho artifact — agent dùng nguyên văn, không phải tự dịch mã kỹ thuật.

    Kỳ tháng trước đã chốt nên "đủ" ở đó nghĩa là đủ CẢ THÁNG, khác nghĩa với kỳ tháng này (đủ tới
    ngày hôm qua) — nói rõ để tin không mơ hồ.
    """
    mx = vr.get("max_ngay")
    if state == cron_status.STATE_DU:
        return (f"đã có số liệu đến ngày {mx[8:10]}/{mx[5:7]}" if (mx and la_ky_chinh)
                else "đủ số liệu cả tháng")
    if state == cron_status.STATE_CHAM:
        if mx:
            tre = (datetime.strptime(ngay_can, "%Y-%m-%d") - datetime.strptime(mx, "%Y-%m-%d")).days
            return (f"file đã kéo về nhưng kế toán chưa nhập số liệu ngày {ngay_can[8:10]}/"
                    f"{ngay_can[5:7]}, số liệu mới nhất là ngày {mx[8:10]}/{mx[5:7]}"
                    f" (chậm {tre} ngày)")
        return "file đã kéo về nhưng chưa có số liệu của ngày cần"
    if state == cron_status.STATE_LOI_NAP:
        return ("hệ thống không đọc được số liệu trong file"
                if vr.get("code") != "KHONG_CO_DATASET"
                else "chưa có kỳ dữ liệu này trên dashboard")
    return "chưa thấy file báo cáo ở nguồn"


if __name__ == "__main__":
    sys.exit(main())

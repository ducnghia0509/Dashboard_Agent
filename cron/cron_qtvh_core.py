# -*- coding: utf-8 -*-
"""ĐỘNG CƠ CHUNG cho 2 job kéo nguồn QTVH: cron_srvf_daily.py / cron_xdv_daily.py.

VÌ SAO TÁCH RA MODULE NÀY (không copy khung của cron_hqkdngay_daily.py lần thứ 3): phần
"xin metadata -> xin file -> chờ về -> autofill -> kiểm -> ghi artifact" giống nhau từng dòng
giữa 3 job; chỉ có DANH SÁCH NGUỒN và CÁCH CHỌN BẢN là khác. Nhân bản khung ra 3 chỗ thì mỗi lần
sửa một cái bẫy (đã có 4 cái, xem dưới) phải nhớ sửa 3 nơi — kiểu lỗi đã xảy ra thật với
`thang_tu_ten_file` (chỉ vá ở chỗ derive, quên chỗ CHỌN file, Dự án rơi khỏi mọi lượt kéo).
Hai job vẫn là 2 SCRIPT, 2 dòng crontab, 2 log, 2 artifact riêng — theo yêu cầu user 24/08/2026.

KHÁC `cron_hqkdngay_daily.py` Ở ĐÂU (và vì sao không dùng lại nó):
job đó phục vụ MỘT report_type "baocaohqkdngay" nơi mỗi đơn vị đúng 1 file/tháng và file là
workbook luỹ kế — nên "chọn file nào" không phải câu hỏi. Nguồn QTVH thì ngược lại: riêng
`SRVF/baocaokqkd` có 15 entry cho kỳ T8 (3 kênh × 5 bản chốt). Chọn sai là cộng đôi doanh thu.

── BỐN CÁI BẪY, ĐỀU ĐÃ CÓ BẰNG CHỨNG TRÊN PROD 24/08/2026 ──────────────────────────────────
1. NHIỀU BẢN CHỐT CÙNG MỘT KỲ -> CỘNG ĐÔI. `Xuathoadon_B2C_T8` có 3 bản (2026.7.4.23 /
   2026.8.3.21 / 2026.8.3.22), mỗi bản là LUỸ KẾ CẢ THÁNG. Lọc theo `month == 8` như 2 job cũ sẽ
   nạp cả 3 -> doanh thu bán xe nhân ba. DB hiện đúng CHỈ VÌ có người nạp tay đúng bản .22.
   -> chế độ `luy_ke`: mỗi (report_type, kỳ, slot) chỉ giữ ĐÚNG 1 bản mới nhất, và XOÁ rows của
      các bản cũ cùng slot (user chốt 24/08: xoá, không dùng hidden_files).
2. NGƯỢC LẠI, ẢNH CHỤP TUẦN PHẢI GIỮ NHIỀU BẢN. 5 nguồn KSCL của XDV nằm trong `_SNAP_RT`, mỗi
   tuần một ngày chốt riêng và màn hình cần cả chuỗi. Xoá bản cũ ở đây là mất dữ liệu thật.
   -> chế độ `anh_chup_ky`: nạp MỌI bản chưa có, KHÔNG xoá gì.
3. `_SNAP_RT` CHỌN NGÀY CHỐT THEO (cong_ty, khoi), KHÔNG THEO source_file. `VHKD_TONVATLY` gán
   cứng `ngay = cuối tháng` cho mọi bản, nên 2 bản khác tên mà cùng `ngay` thì CẢ HAI được chọn và
   cộng vào nhau — `_per_file_resolved` không đỡ được (nó dedupe TRONG một file, không phải giữa
   các file). Đây là lý do (1) phải xoá thật chứ không chỉ "ưu tiên bản mới".
4. FILE DỰNG SẴN CẢ THÁNG -> "có dòng của ngày hôm qua" KHÔNG chứng minh kế toán đã nhập. Kế thừa
   nguyên VÂN TAY (số dòng : Σ|amount|) của cron_hqkdngay_daily.py; nguồn ảnh chụp thì vân tay là
   bằng chứng DUY NHẤT, vì chúng không có trục ngày để so.

XẾP HẠNG "BẢN NÀO MỚI HƠN" = NGÀY CHỐT TRƯỚC, GIỜ SỬA SAU (xem `_xep_slot`). Ngày chốt đọc từ tên
file qua `ngay_regex` KHAI RIÊNG TỪNG THƯ MỤC — không có công thức chung vì tên file có ít nhất 3
dạng token ('2026.8.3.22' = năm.tháng.tuần.ngày, '2026.8.24' = năm.tháng.ngày, '2026.8.17.6' =
không đọc được chắc chắn). Thư mục nào dạng tên còn nhập nhằng thì BỎ TRỐNG `ngay_regex` và cả slot
lùi về `(modifiedAt, createdAt, fileName)`.

Vì sao không xếp thuần `modifiedAt` cho gọn (dry-run 24/08/2026 bắt được): kho xe B2B có bản chốt
17/08 bị sửa lại lúc 13:33 hôm nay, muộn hơn bản chốt 24/08 (09:54) — xếp theo giờ sửa là giữ ảnh
chụp ngày 17 và XOÁ ảnh chụp ngày 24. Với ảnh chụp, "mới hơn" là chốt muộn hơn, không phải được gõ
muộn hơn. Log luôn in cả ngày chốt và giờ sửa của bản giữ + mọi bản loại để người soi lại được.

KỲ CỦA DỮ LIỆU CŨNG PHẢI ĐỌC TỪ TÊN FILE (`ky_regex`), KHÔNG TIN `month` CỦA METADATA — xem
`ky_cua`. Hai ca đã bắt được: `Tonkhoxevatly_T1_B2B` bị gán month=7, và `BaocaoClaim_B2C_T11.25`
(tháng 11 năm 2025) cũng bị gán month=7 nên suýt bị xoá như bản chốt cũ của kỳ 2026-07.

Không phải đường ra nào cũng xoá: THỨ TỰ BẮT BUỘC là autofill bản mới -> kiểm có dòng thật -> MỚI
xoá bản cũ. Xoá trước rồi nạp lỗi là mất cả hai.
"""
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cron_status                                               # noqa: E402

VN = timezone(timedelta(hours=7))
ROOT = "/home/itadmin/AI_Dashboard_QT"
CONNECT = f"{ROOT}/Connect_VPS"
AVAILABLE_META = f"{CONNECT}/available_metadata.json"
RECEIVED_DIR = f"{CONNECT}/received_reports"
AGENT = f"{ROOT}/Dashboard_Agent"
AGENT_PY = f"{AGENT}/.venv/bin/python"
RECEIVER_URL = "http://127.0.0.1:8090"

META_REFRESH_TIMEOUT = 240
META_REFRESH_POLL = 5
ARRIVE_TIMEOUT = 300
ARRIVE_POLL = 5

# 3 chế độ — xem khối bẫy (1)(2) ở docstring. Đặt thành hằng chuỗi để bảng khai báo của 2 job đọc
# được như tài liệu, và để gõ sai tên chế độ thì nổ ngay lúc nạp module chứ không im lặng.
LUY_KE = "luy_ke"              # luỹ kế cả kỳ, 1 bản/slot, xoá bản cũ
ANH_CHUP_KY = "anh_chup_ky"    # ảnh chụp từng kỳ (tuần), giữ mọi bản
THANG = "thang"                # 1 file/tháng, kéo lại đè chính nó
CHE_DO = (LUY_KE, ANH_CHUP_KY, THANG)

# `anh_chup_ky`: số bản chốt mới nhất mỗi slot LUÔN nạp lại, kể cả khi DB đã có. Xem chú thích ở
# `pick_targets`. Để 1 là kế toán sửa lại file tuần trước sẽ không bao giờ được nạp lại.
GIU_MOI_NHAT = 2


def moi_truong(job: str) -> dict:
    """ENVIRONMENTS của job — cùng quy ước đường dẫn/hậu tố với 2 job cũ.

    Log + artifact ghi vào `AI_coding/logs/` (KHÔNG phải cạnh script): hằng số `_LOGS` của
    source_bridge.py (panel "Nguồn dữ liệu" > Chạy tự động) và agent giám sát trên openclaw đều
    trỏ vào đó. Xem docstring cron_hqkdngay_daily.py, cùng lý do.
    """
    logs = f"{ROOT}/AI_coding/logs"
    return {
        "test": {
            "database_url": "postgresql://tc:tc_%24production@localhost:5435/tc_dashboard",
            "log_file": f"{logs}/cron_{job}.log",
            "disabled_flag": f"{logs}/cron_{job}.disabled",
            "verify_api_dir": f"{ROOT}/AI_coding/tc-admin-api",
            "status_json": f"{logs}/status_{job}.json",
            "status_jsonl": f"{logs}/status_{job}.jsonl",
        },
        "prod": {
            "database_url": "postgresql://tc:tc_%24production@localhost:5434/tc_dashboard",
            "log_file": f"{logs}/cron_{job}_prod.log",
            "disabled_flag": f"{logs}/cron_{job}_prod.disabled",
            "verify_api_dir": "/home/itadmin/apps/tc-console/tc-admin-api",
            "status_json": f"{logs}/status_{job}_prod.json",
            "status_jsonl": f"{logs}/status_{job}_prod.jsonl",
        },
    }


class Ctx:
    """Trạng thái một lượt chạy — thay biến global của 2 job cũ.

    2 job cũ dùng `global LOG_FILE, DATABASE_URL, ...` rồi gán lại trong main(). Ở đây engine bị
    2 script gọi nên global là mời race khi có ngày nào đó ai chạy 2 job trong một tiến trình.
    """

    def __init__(self, job: str, env: str, cfg: dict, dry_run: bool):
        self.job, self.env, self.cfg, self.dry_run = job, env, cfg, dry_run
        self.log_file = cfg["log_file"]
        if dry_run and self.log_file.endswith(".log"):
            # DRY-RUN ghi log SANG FILE RIÊNG: log chính là nguồn duy nhất của panel "Chạy tự động
            # theo lịch" (source_bridge._job_last_run đọc khối cuối) — một lượt dry-run vào log
            # chính làm job hiện ĐỎ OAN tới lần chạy thật kế tiếp (gặp thật 12/08/2026).
            self.log_file = self.log_file[:-4] + "_dryrun.log"

    def log(self, msg: str):
        line = f"[{datetime.now(VN):%Y-%m-%d %H:%M:%S} VN] {msg}"
        print(line, flush=True)
        os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
        with open(self.log_file, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")


# ── receiver ────────────────────────────────────────────────────────────────────────────────
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


def wait_metadata(ctx: Ctx, before) -> bool:
    t0 = time.time()
    while time.time() - t0 < META_REFRESH_TIMEOUT:
        time.sleep(META_REFRESH_POLL)
        if meta_mtime() != before:
            time.sleep(1)
            ctx.log(f"  danh sách mới về sau {int(time.time() - t0)}s")
            return True
    ctx.log(f"  CANH BAO: {META_REFRESH_TIMEOUT}s máy Local chưa nộp danh sách mới (agent còn sống"
            " không?) -> dùng danh sách CŨ, mốc giờ dưới đây là của lần quét trước")
    return False


def read_metadata(ctx: Ctx):
    for _ in range(3):
        try:
            with open(AVAILABLE_META, encoding="utf-8") as fh:
                return json.load(fh)
        except json.JSONDecodeError:
            time.sleep(2)
        except OSError as ex:
            ctx.log(f"DỪNG: không đọc được {AVAILABLE_META} ({ex})")
            return None
    ctx.log(f"DỪNG: {AVAILABLE_META} đọc 3 lần vẫn lỗi JSON (ghi dở hoặc hỏng)")
    return None


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
    """'<thư_mục_công_ty>::<tên_file>' — khớp `source_file` trong raw_rows (spec_extract._source_id
    và source_catalog.source_id_from_path cùng quy ước)."""
    return f"{entry.get('company') or ''}::{entry['fileName']}"


# ── chọn bản ────────────────────────────────────────────────────────────────────────────────
def target_periods() -> list:
    """[(period, month, year)] = tháng NÀY + tháng TRƯỚC theo giờ VN (đầu tháng còn chốt sổ cũ)."""
    now = datetime.now(VN)
    prev = now.replace(day=1) - timedelta(days=1)
    return [(f"{d.year}-{d.month:02d}", d.month, d.year) for d in (now, prev)]


def ngay_chot(nguon: dict, e: dict):
    """'YYYY-MM-DD' NGÀY CHỐT đọc từ tên file, hoặc None nếu thư mục không khai `ngay_regex`.

    `ngay_regex` phải bắt 3 nhóm (năm, tháng, ngày). Chỉ khai cho thư mục có dạng tên ĐỌC ĐƯỢC
    CHẮC CHẮN; dạng nào còn nhập nhằng thì bỏ trống và để xếp theo `modifiedAt`.
    """
    pat = nguon.get("ngay_regex")
    if not pat:
        return None
    m = re.search(pat, e.get("fileName") or "", re.IGNORECASE)
    if not m:
        return None
    try:
        y, mo, d = (int(x) for x in m.group(1, 2, 3))
        return f"{y:04d}-{mo:02d}-{d:02d}" if 1 <= mo <= 12 and 1 <= d <= 31 else None
    except (ValueError, IndexError):
        return None


def _vintage_key(e: dict) -> tuple:
    """Khoá xếp hạng "bản nào mới hơn" trong một slot — xem `_xep_slot`."""
    return (e.get("modifiedAt") or "", e.get("createdAt") or "", e.get("fileName") or "")


def _xep_slot(ds: list) -> list:
    """Xếp các bản chốt của MỘT slot từ cũ -> mới. NGÀY CHỐT thắng GIỜ SỬA.

    Vì sao không xếp thuần `modifiedAt` (bản đầu làm thế, dry-run 24/08/2026 bắt được): kho xe B2B
    có `2026.8.17.KHO XE` bị sửa lại lúc 13:33 hôm nay, muộn hơn `2026.8.24.KHO XE` (09:54) — xếp
    theo giờ sửa là chọn ảnh chụp ngày 17 và XOÁ ảnh chụp ngày 24. Với nguồn ảnh chụp, "mới hơn"
    nghĩa là chốt ở thời điểm muộn hơn, không phải được gõ vào muộn hơn.

    Chỉ dùng ngày chốt khi MỌI bản trong slot đọc được nó; sót một bản là bản đó có khoá rỗng và
    tự động xuống bét — đúng cái bẫy vừa tránh, nên cả slot lùi về `modifiedAt`.
    """
    ngay = [ngay_chot(e["_nguon"], e) for e in ds]
    if all(ngay):
        return [e for _, e in sorted(zip(ngay, ds), key=lambda t: (t[0], _vintage_key(t[1])))]
    return sorted(ds, key=_vintage_key)


def ky_cua(ctx: Ctx, nguon: dict, e: dict):
    """(thang, nam) THẬT của dữ liệu trong file — `ky_regex` thắng `month` của metadata.

    VÌ SAO KHÔNG TIN `month` CỦA METADATA cho nhóm SRVF (bắt được bằng dry-run 24/08/2026): bên
    quét bóc tháng từ hậu tố `_T{n}` khi nó nằm CUỐI tên ('Xuathoadon_B2B_T8' -> month=8 ✓) nhưng
    rơi về token ngày tạo khi `_T{n}` nằm GIỮA ('Tonkhoxevatly_T1_B2B' -> month=7 ✗, đúng phải là
    1). Hậu quả nếu tin nó: kỳ 2026-07 gom cả 22 file tồn vật lý T1..T7 vào một lượt -> mỗi ngày
    nạp lại 6 tháng lịch sử, và chỉ thoát cộng đôi nhờ `slot` có chứa số tháng.

    Năm lấy từ token 20xx của tên file khi tên chỉ có DUY NHẤT một năm; nhiều/không có thì trả None
    (bên gọi bỏ qua phép kiểm năm, y như đường cũ).
    """
    pat = nguon.get("ky_regex")
    fn = e.get("fileName") or ""
    if pat:
        m = re.search(pat, fn, re.IGNORECASE)
        if m:
            nam = {int(t) for t in re.findall(r"(20\d{2})", fn)}
            return int(m.group(1)), (nam.pop() if len(nam) == 1 else None)
        # NỔ TO thay vì âm thầm rơi về metadata: khai `ky_regex` nghĩa là "tên file thư mục này có
        # mang kỳ", không khớp là đã có dạng tên mới cần bổ sung regex.
        ctx.log(f"  CANH BAO [{nguon['company']}/{nguon['rt']}]: ten file '{fn[:52]}' khong khop"
                f" ky_regex -> tam dung 'month' cua metadata ({e.get('month')})")
    return cron_status.ky_cua_entry(e)


def _slot(nguon: dict, e: dict, period: str) -> tuple:
    """Định danh "cùng một thứ, khác bản chốt".

    Mặc định (report_type, kỳ) — đúng cho mọi thư mục 1 báo cáo/kỳ. Thư mục nào một kỳ có nhiều
    LÁT song song (kênh B2B/B2C/GF) thì khai `slot` = regex, các nhóm bắt được ghép vào khoá; thiếu
    nó thì 3 kênh bị coi là 3 bản chốt của nhau và 2 kênh bị xoá oan.
    """
    extra = ()
    pat = nguon.get("slot")
    if pat:
        m = re.search(pat, e.get("fileName") or "", re.IGNORECASE)
        # Không khớp -> slot rỗng (đứng riêng một mình), KHÔNG gộp vào slot của bản khác: tên file
        # lạ mà bị gộp là xoá mất một lát dữ liệu thật.
        extra = tuple(x or "" for x in m.groups()) if m else (e.get("fileName") or "",)
    return (nguon["rt"], period) + tuple(s.upper() for s in extra)


def pick_targets(ctx: Ctx, nguon_list: list, meta: list, periods: list):
    """Trả (targets, losers). `targets` = bản sẽ nạp; `losers` = bản cũ cùng slot sẽ xoá rows.

    Chỉ chế độ `luy_ke` sinh `losers`. `anh_chup_ky` giữ hết, `thang` mỗi kỳ vốn 1 file.
    """
    by_key = {}
    for nguon in nguon_list:
        for period, month, year in periods:
            for e in meta:
                if e.get("company") != nguon["company"] or e.get("report_type") != nguon["rt"]:
                    continue
                if not e.get("fileName"):
                    continue
                fn = e["fileName"]
                # `chi_lay`/`bo_qua`: một thư mục có thể chứa loại báo cáo mà KHÔNG spec nào bóc
                # (SRVF/baocaokqkd còn 'Kymoi_*' — hợp đồng ký mới, chưa có spec). Kéo về là đĩa
                # thêm file rác và log thêm cảnh báo giả "nạp 0 dòng".
                if nguon.get("chi_lay") and not re.search(nguon["chi_lay"], fn, re.IGNORECASE):
                    continue
                if nguon.get("bo_qua") and re.search(nguon["bo_qua"], fn, re.IGNORECASE):
                    continue
                thang, nam_ten = ky_cua(ctx, nguon, e)
                if thang != month:
                    continue
                years = {int(t) for t in re.findall(r"(20\d{2})", fn)}
                if years and year not in years:
                    continue
                if nam_ten is not None and nam_ten != year:
                    continue
                by_key.setdefault(_slot(nguon, e, period), []).append(
                    {**e, "_period": period, "_nguon": nguon})

    targets, losers = [], []
    for key, ds in sorted(by_key.items()):
        mode = ds[0]["_nguon"]["che_do"]
        if mode == ANH_CHUP_KY:
            # Ảnh chụp tuần đã nạp rồi thì nạp lại KHÔNG đổi gì (idempotent theo source_file) —
            # nhưng vẫn tốn 1 lượt xin file + 1 lượt derive mỗi ngày cho MỌI tuần của 2 kỳ đang
            # xem (đo được: 23 file/lượt bên XDV). Đánh dấu các bản CŨ hơn 2 bản mới nhất là "bỏ
            # qua được"; `run()` sẽ bỏ đúng những bản đã có dòng trong DB. Giữ 2 bản (không phải 1)
            # để kế toán sửa lại file tuần trước thì lượt sau vẫn nạp lại.
            ds = _xep_slot(ds)
            for e in ds[:-GIU_MOI_NHAT]:
                e["_bo_qua_neu_co_roi"] = True
            targets.extend(ds)
            continue
        if mode != LUY_KE or len(ds) == 1:
            targets.extend(ds)
            continue
        ds = _xep_slot(ds)
        win, old = ds[-1], ds[:-1]
        targets.append(win)
        losers.extend(old)
        ctx.log(f"  TRÙNG KỲ [{key[0]} {key[1]}{' ' + '/'.join(key[2:]) if len(key) > 2 else ''}]:"
                f" {len(ds)} bản -> giữ {win['fileName'][:46]}"
                f" (chốt {ngay_chot(win['_nguon'], win) or '?'},"
                f" sửa {win.get('modifiedAt')})")
        for o in old:
            ctx.log(f"      loại {o['fileName'][:46]} (chốt {ngay_chot(o['_nguon'], o) or '?'},"
                    f" sửa {o.get('modifiedAt')})")
    return targets, losers


# ── kéo file ────────────────────────────────────────────────────────────────────────────────
def request_files(ctx: Ctx, targets: list) -> dict:
    before = {}
    for e in targets:
        before[e["fileName"]] = sidecar_saved_at(e)
        payload = {k: e.get(k) for k in ("company", "report_type", "fileName", "path",
                                         "month", "periodType", "status")}
        try:
            http("POST", "/request-file", payload)
            ctx.log(f"  đã xin: [{e.get('company')}/{e.get('report_type')}] {e['fileName'][:50]}")
        except (urllib.error.URLError, OSError) as ex:
            ctx.log(f"  LỖI xin file {e['fileName'][:46]}: {ex}")
    return before


def wait_arrival(ctx: Ctx, targets: list, before: dict) -> set:
    deadline = time.time() + ARRIVE_TIMEOUT
    pending = {e["fileName"]: e for e in targets}
    arrived = set()
    while pending and time.time() < deadline:
        time.sleep(ARRIVE_POLL)
        for fn, e in list(pending.items()):
            now_at = sidecar_saved_at(e)
            if now_at and now_at != before.get(fn):
                arrived.add(fn)
                del pending[fn]
                ctx.log(f"  đã về: {fn[:52]} (saved_at {now_at})")
    for fn, e in pending.items():
        old = before.get(fn)
        ctx.log(f"  KHÔNG VỀ sau {ARRIVE_TIMEOUT}s: {fn[:48]}"
                + (f" — vẫn dùng bản cũ trên đĩa {old}" if old else " — trên đĩa chưa có bản nào"))
    return arrived


def autofill(ctx: Ctx, entry: dict):
    """`agent_cli.py autofill` — điểm vào DUY NHẤT, tự dispatch: báo cáo ngày -> derive_hqkd_ngay,
    nguồn khai bằng spec JSON (toàn bộ VHKD + XDV) -> spec_extract, còn lại -> đường tất định.
    Có khoá per-file (servers/common/filelock.py) nên 2 job này chạy lệch giờ vẫn an toàn nếu
    trùng, và an toàn cả khi user bấm "Chạy ngay" trên panel Nguồn dữ liệu cùng lúc.

    Trả (ok, ly_do, canh_bao). rc=0 mà JSON `ok:false` VẪN là lỗi nạp — agent_cli không đổi exit
    code cho ca đó.
    """
    path = xlsx_path(entry)
    env = {**os.environ, "DATABASE_URL": ctx.cfg["database_url"]}
    cmd = [AGENT_PY, "scripts/agent_cli.py", "autofill", path, "--period", entry["_period"]]
    p = subprocess.run(cmd, cwd=AGENT, env=env, capture_output=True, text=True, timeout=900)
    tail = (p.stdout or p.stderr or "").strip().splitlines()
    last = tail[-1] if tail else ""
    ctx.log(f"  autofill [{entry.get('report_type')}] {entry['_period']}: rc={p.returncode}"
            f" | {last[:260]}")
    if p.returncode != 0:
        return False, f"autofill lỗi (rc={p.returncode})", []
    try:
        js = json.loads(last)
    except (json.JSONDecodeError, ValueError):
        return True, None, []            # không parse được thì để verify quyết định
    cb = list(js.get("canh_bao") or [])
    # KỲ MÀ FILE THỰC SỰ GHI VÀO có thể KHÁC kỳ suy từ tên file — và đó là một lỗi dữ liệu thật,
    # không phải chuyện lý thuyết. Prod 24/08/2026: `Xuathoadon_GF_T7.xlsx` chứa toàn số THÁNG 6
    # nên ghi vào kỳ 2026-06, trùng khít 11 dòng / 5,533 tỷ với `Xuathoadon_GF_T6.xlsx` -> GF
    # tháng 6 cộng đôi, mà `luy_ke` KHÔNG đỡ được (hai file khác slot theo tên nên đều là bản chốt
    # hợp lệ). Cron không tự sửa được ca này (lỗi ở NỘI DUNG file, kế toán phải sửa), nhưng phải
    # NÓI TO — nếu không nó lại nằm im vài tháng như lần này.
    ky_ghi = {k for d in (js.get("derived") or []) for k in (d.get("ky") or {})}
    lech = sorted(k for k in ky_ghi if k != entry["_period"])
    if lech:
        cb.append(f"KỲ LỆCH: tên file thuộc kỳ {entry['_period']} nhưng nội dung ghi vào kỳ "
                  f"{', '.join(lech)} — kiểm tra file nguồn có bị dán nhãn sai tháng không "
                  f"(có thể đang cộng đôi với file của kỳ đó)")
    if js.get("skipped_lock"):
        # Bị khoá = lượt khác ĐANG nạp đúng file này. Không phải lỗi file, và tuyệt đối không được
        # coi là "nạp xong" rồi đi xoá bản cũ.
        return False, "file đang được một lượt nạp khác xử lý (khoá per-file)", cb
    if js.get("ok"):
        return True, None, cb
    errs = [d.get("error") for d in (js.get("derived") or []) if d.get("error")]
    return False, (errs[0][:200] if errs else "autofill trả ok=false, không kèm lý do"), cb


# ── kiểm + xoá bản cũ (cùng một đường vào DB) ───────────────────────────────────────────────
def _py_sql(ctx: Ctx, code: str, timeout=180) -> str:
    """Chạy 1 đoạn python trong venv của API để dùng `app.database.session` — cùng cách
    cron_hqkdngay_daily.verify() làm. Cố ý KHÔNG mở kết nối psycopg riêng ở đây: DSN, pool và
    lớp `?`->`%s` của dự án nằm trong session đó, dựng lại là thêm một nguồn sự thật thứ hai."""
    api = ctx.cfg["verify_api_dir"]
    p = subprocess.run([f"{api}/.venv/bin/python", "-c", code], cwd=api,
                       env={**os.environ, "DATABASE_URL": ctx.cfg["database_url"]},
                       capture_output=True, text=True, timeout=timeout)
    return (p.stdout or p.stderr).strip()


def verify(ctx: Ctx, entry: dict) -> dict:
    """Kiểm theo ĐÚNG source_file: có dòng thật không, trải mấy ngày, và VÂN TAY của file.

    KHÔNG đòi "max(ngay) >= hôm qua" như job báo cáo ngày: phần lớn nguồn ở đây là ẢNH CHỤP, `ngay`
    của chúng là ngày chốt (cuối tháng, hoặc Chủ nhật của tuần) chứ không phải ngày phát sinh — đòi
    ngày hôm qua là báo chậm oan mỗi lượt. Bằng chứng "kế toán đã cập nhật" ở đây là VÂN TAY ĐỔI.
    """
    sid = source_id(entry)
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.database.session import get_db;"
        f"sid={sid!r};"
        "r=get_db().execute(\"SELECT COUNT(DISTINCT ngay) nd,MAX(ngay) mx,COUNT(*) sd,"
        "SUM(ABS(COALESCE(amount,0))) tg,COUNT(DISTINCT report_type) nrt FROM raw_rows "
        "WHERE source_file=?\",(sid,)).fetchone();"
        "print('KHONG_CO_DONG_NAO') if not r or not r['sd'] else "
        "print('so_ngay=%s max_ngay=%s so_rt=%s van_tay=%s:%.6f CO_DONG'%(r['nd'],r['mx'],"
        "r['nrt'],r['sd'],float(r['tg'] or 0)))"
    )
    out = _py_sql(ctx, code)
    ctx.log(f"  kiem [{entry.get('report_type')}] {entry['_period']}: {out[:200]}")
    return parse_verify(out)


def parse_verify(out: str) -> dict:
    d = {"raw": out[:200], "so_ngay": None, "max_ngay": None, "code": None}
    for code in ("KHONG_CO_DONG_NAO", "CO_DONG"):
        if code in out:
            d["code"] = code
            break
    for k, pat in (("so_ngay", r"so_ngay=(\d+)"), ("max_ngay", r"max_ngay=(\d{4}-\d{2}-\d{2})"),
                   ("van_tay", r"van_tay=(\d+:[\d.]+)")):
        m = re.search(pat, out)
        if m:
            d[k] = int(m.group(1)) if k == "so_ngay" else m.group(1)
    return d


def da_co_dong(ctx: Ctx, sids: list) -> set:
    """Tập `source_file` ĐÃ có dòng trong DB — dùng để bỏ qua ảnh chụp tuần cũ đã nạp."""
    if not sids:
        return set()
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.database.session import get_db;"
        f"sids={sorted(set(sids))!r};"
        "\nfor s in sids:\n"
        "    r=get_db().execute('SELECT COUNT(*) c FROM raw_rows WHERE source_file=?',"
        "(s,)).fetchone()\n"
        "    if (r['c'] if r else 0):\n"
        "        print('CO|%s'%s)\n"
    )
    out = _py_sql(ctx, code, timeout=300)
    return {ln[3:] for ln in out.splitlines() if ln.startswith("CO|")}


def trang_thai(vr: dict, van_tay_cu, doi_luc_cu, today: str):
    """(state, doi_luc) cho nguồn ẢNH CHỤP — bản song song của `cron_status.state_from_verify`.

    KHÔNG dùng lại hàm đó: nó phân loại theo `OK_CO_NGAY_HOM_QUA`/`THIEU_NGAY_HOM_QUA`, tức giả
    định nguồn có trục ngày phát sinh. Nguồn ở đây phần lớn là ảnh chụp (ngày = ngày chốt), nên
    tiêu chí duy nhất còn ý nghĩa là VÂN TAY ĐỔI. Đặt cạnh nhau ở 2 hàm rõ ràng hơn là nhồi cả 2
    tiêu chí vào một hàm rồi truyền cờ để chọn nhánh.

    Lượt ĐẦU (chưa có vân tay cũ) không kết luận được "kế toán đã cập nhật hay chưa" -> xếp `du`
    (có số thật, không báo động), và từ lượt sau mới so được.
    """
    doi_luc = today if (vr.get("van_tay") and van_tay_cu
                        and vr["van_tay"] != van_tay_cu) else doi_luc_cu
    if vr.get("code") != "CO_DONG":
        return cron_status.STATE_LOI_NAP, doi_luc
    if not van_tay_cu:
        return cron_status.STATE_DU, doi_luc
    return (cron_status.STATE_DU if doi_luc == today else cron_status.STATE_CHAM), doi_luc


def xoa_ban_cu(ctx: Ctx, losers: list, da_nap_ok: set) -> int:
    """Xoá rows của các bản chốt CŨ cùng slot — chỉ khi bản MỚI của slot đó đã nạp thành công.

    Vì sao xoá thật (user chốt 24/08/2026) chứ không ghi `hidden_files`: `_SNAP_RT` chọn ngày chốt
    theo (cong_ty, khoi) nên 2 bản cùng `ngay` vẫn được cộng cả hai — ẩn ở tầng hiển thị không cứu
    được các đường cộng khác, và mọi truy vấn psql ad-hoc vẫn thấy số cộng đôi.

    Xoá theo `source_file` TRỌN GÓI (không kèm report_type): một file cấp nhiều report_type
    (baocaocongnophaithu -> VHKD_PTHU + _KENH + _COC), sót loại nào là còn đúng cái cộng đôi đó.
    """
    can_xoa = [o for o in losers if o["_slot_key"] in da_nap_ok]
    bo_qua = [o for o in losers if o["_slot_key"] not in da_nap_ok]
    for o in bo_qua:
        ctx.log(f"  GIỮ NGUYÊN bản cũ {o['fileName'][:46]}: bản mới cùng slot chưa nạp được"
                " -> xoá bây giờ là mất cả hai")
    if not can_xoa:
        return 0
    sids = sorted({source_id(o) for o in can_xoa})
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.database.session import get_db;"
        f"sids={sids!r};db=get_db();n=0;"
        "\nfor s in sids:\n"
        "    r=db.execute('SELECT COUNT(*) c FROM raw_rows WHERE source_file=?',(s,)).fetchone()\n"
        "    c=(r['c'] if r else 0) or 0\n"
        "    if c:\n"
        "        db.execute('DELETE FROM raw_rows WHERE source_file=?',(s,))\n"
        "    n+=c\n"
        "    print('XOA %s dong | %s'%(c,s))\n"
        "print('TONG_XOA=%s'%n)"
    )
    out = _py_sql(ctx, code, timeout=300)
    for line in out.splitlines():
        ctx.log(f"  {line[:180]}")
    m = re.search(r"TONG_XOA=(\d+)", out)
    return int(m.group(1)) if m else 0


# ── rà soát sau khi nạp ─────────────────────────────────────────────────────────────────────
# Hai lớp lỗi dưới đây KHÔNG làm job đỏ (file vẫn về, dòng vẫn nạp) nhưng làm SỐ TRÊN DASHBOARD
# SAI, và trước 25/08/2026 chúng chỉ lộ ra khi có người ngồi truy vấn tay. Bằng chứng: prod cõng
# `Tonkhoxevatly_T7.xlsx` (bản gộp) cạnh 3 file kênh T7 -> 852 VIN đếm hai lần, im lặng cả tháng.
def _ten_goc(fn: str) -> str:
    """Tên báo cáo sau khi bỏ TIỀN TỐ NGÀY PHÁT HÀNH ('B.1.TC.OO.M.2026.7.4.23.').

    Đây là khoá phân biệt hai thứ nhìn na ná nhau trong cùng một ngày chốt:
    · hai LÁT SONG SONG   -> tên gốc KHÁC nhau (Tonkhoxevatly_T8_B2B vs _B2C) — bình thường;
    · hai BẢN CHỐT của cùng một báo cáo -> tên gốc GIỐNG nhau (bản 14/08 vs bản 21/08) — cộng đôi.
    """
    s = fn.split("::", 1)[-1]
    s = re.sub(r"^.*?\b20\d{2}(?:[.\-_]\d{1,2}){0,4}[.\-_]", "", s)   # bỏ tiền tố ngày phát hành
    s = re.sub(r"^\d{1,2}[.\-_]", "", s)                              # sót token ngày lẻ
    return s.strip().lower()


def _la_ban_gop(a: str, b: str) -> bool:
    """`a` là BẢN GỘP của `b`: cùng đuôi, tên của a là tiền tố của b và cắt ĐÚNG ở dấu phân tách.

    'Tonkhoxevatly_T7' ⊃ 'Tonkhoxevatly_T7_B2B' (cắt ở '_') -> True. Đòi cắt đúng dấu phân tách
    để 'Tonkhoxevatly_T7' vs 'Tonkhoxevatly_T71' KHÔNG thành báo động giả.
    """
    ta, ea = os.path.splitext(a)
    tb, eb = os.path.splitext(b)
    return ea == eb and len(tb) > len(ta) and tb.startswith(ta) and tb[len(ta)] in "_-. "


def _sids_tren_dia(nguon_list: list) -> list:
    """[(source_id, nguon, fileName)] cho MỌI file Excel đang nằm trong thư mục của các nguồn khai
    báo — qua đúng bộ lọc `chi_lay`/`bo_qua` của nguồn đó, để file CỐ Ý không kéo (Kymoi_*, bản
    gộp không hậu tố kênh) không bị báo là mồ côi."""
    out = []
    for nguon in nguon_list:
        d = os.path.join(RECEIVED_DIR, nguon["company"], nguon["rt"])
        try:
            names = sorted(os.listdir(d))
        except OSError:
            continue
        for fn in names:
            if not fn.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".xls")) or fn.startswith("~$"):
                continue
            if nguon.get("chi_lay") and not re.search(nguon["chi_lay"], fn, re.IGNORECASE):
                continue
            if nguon.get("bo_qua") and re.search(nguon["bo_qua"], fn, re.IGNORECASE):
                continue
            out.append((f"{nguon['company']}::{fn}", nguon, fn))
    return out


def _ban_moi_nhat_moi_slot(ctx: Ctx, tren_dia: list) -> list:
    """Lọc `tren_dia` xuống còn BẢN MỚI NHẤT của mỗi slot — dùng lại đúng `_slot`/`_xep_slot` của
    đường kéo, để "mồ côi" mang cùng một định nghĩa với "bản cron sẽ giữ".

    Nếu không lọc: mọi bản chốt CŨ mà `xoa_ban_cu` vừa dọn rows đều hiện ra như file mồ côi (đo
    trên prod 25/08: 20/36 báo động là loại này). Ngày chốt đọc từ tên file, thiếu thì lùi về
    mtime của chính file trên đĩa — cùng thứ tự ưu tiên với `_xep_slot`.
    """
    nhom = {}
    for sid, nguon, fn in tren_dia:
        e = {"fileName": fn, "_nguon": nguon}
        try:
            e["modifiedAt"] = datetime.fromtimestamp(
                os.path.getmtime(os.path.join(RECEIVED_DIR, nguon["company"], nguon["rt"], fn)),
                VN).strftime("%Y-%m-%d %H:%M:%S")
        except OSError:
            e["modifiedAt"] = ""
        thang, _ = ky_cua(ctx, nguon, e)
        e["_ky"] = f"T{thang}" if thang else "?"
        if nguon["che_do"] == ANH_CHUP_KY:
            # Ảnh chụp tuần: MỌI bản đều phải có trong DB (cả chuỗi tuần mới vẽ được biểu đồ) nên
            # mỗi file là một slot riêng — gom lại rồi chỉ soi bản mới nhất là bỏ lọt đúng ca hay
            # hỏng nhất: một tuần lẻ giữa chuỗi nạp trượt mà không ai biết. Nhãn kỳ lấy NGÀY CHỐT
            # vì cả thư mục cùng một tháng.
            e["_ky"] = ngay_chot(nguon, e) or e["_ky"]
            nhom[(nguon["rt"], fn)] = [(sid, nguon, fn, e)]
            continue
        nhom.setdefault(_slot(nguon, e, e["_ky"]), []).append((sid, nguon, fn, e))
    giu = []
    for _, ds in sorted(nhom.items()):
        moi_nhat = _xep_slot([e for *_, e in ds])[-1]
        giu.extend((sid, nguon, fn, e) for sid, nguon, fn, e in ds if e is moi_nhat)
    return giu


def ra_soat_mo_coi(ctx: Ctx, nguon_list: list) -> list:
    """FILE MỒ CÔI = bản MỚI NHẤT của một slot, có trên đĩa, mà KHÔNG có lấy một dòng nào trong DB.

    Vì sao cần: cửa sổ kéo chỉ gồm tháng này + tháng trước, còn receiver thì chỉ lưu file chứ
    KHÔNG nạp (`receiver_server.upload` ghi file + sidecar rồi dừng). File kỳ cũ — kế toán gửi bù,
    hoặc người ta bấm kéo tay từ giao diện — rơi đúng vào vùng chết giữa hai cái đó: nằm trên đĩa
    mãi mà không ai biết là chưa vào DB.

    CHỈ BÁO, KHÔNG TỰ NẠP: nạp một kỳ cũ có thể phải xoá bản chốt cũ cùng slot, đó là quyết định
    của người chứ không phải của một lượt cron chạy lúc 4 giờ sáng.
    """
    tren_dia = _sids_tren_dia(nguon_list)
    if not tren_dia:
        return []
    ung_vien = _ban_moi_nhat_moi_slot(ctx, tren_dia)
    # Hỏi DB về MỌI file trong thư mục chứ không chỉ các bản mới nhất: phép loại "cùng tên khác
    # đuôi" ở dưới cần biết bản .xlsx (không phải bản mới nhất) có dòng hay không.
    sids = sorted({sid for sid, *_ in tren_dia})
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.database.session import get_db;"
        f"sids={sids!r};"
        "ph=','.join(['?']*len(sids));"
        "\nfor r in get_db().execute('SELECT DISTINCT source_file s FROM raw_rows WHERE "
        "source_file IN (%s)'%ph, tuple(sids)):\n"
        "    print('CO|%s'%r['s'])\n"
    )
    da_co = {ln[3:] for ln in _py_sql(ctx, code, timeout=300).splitlines() if ln.startswith("CO|")}
    # MỘT BÁO CÁO, HAI ĐỊNH DẠNG: thư mục công nợ phải thu giữ song song `..._T8.Xls` (2,2 MB, bản
    # gốc) và `..._T8.xlsx` (126 KB) của cùng một kỳ; chỉ bản .xlsx được nạp. Cùng phần tên, khác
    # đuôi, một bản đã có dòng -> KHÔNG phải mồ côi, đừng bắt người đi tìm một file không thiếu.
    goc_da_co = {os.path.splitext(s)[0].lower() for s in da_co}
    thieu = [(sid, nguon, fn, e) for sid, nguon, fn, e in ung_vien
             if sid not in da_co and os.path.splitext(sid)[0].lower() not in goc_da_co]
    if not thieu:
        ctx.log(f"RÀ SOÁT MỒ CÔI: {len(ung_vien)} bản mới nhất trên đĩa ({len(tren_dia)} file kể"
                " cả bản chốt cũ), tất cả đều đã có dòng trong DB")
        return []
    theo_rt = {}
    for _, nguon, fn, e in thieu:
        theo_rt.setdefault(f"{nguon['company']}/{nguon['rt']}", []).append((fn, e["_ky"]))
    ctx.log(f"RÀ SOÁT MỒ CÔI: {len(thieu)}/{len(ung_vien)} bản mới nhất có trên đĩa nhưng KHÔNG"
            " có dòng nào trong DB (ngoài cửa sổ 2 kỳ của cron, mà receiver thì không nạp)")
    canh_bao = []
    for muc, ds in sorted(theo_rt.items()):
        # Chỉ in THÁNG chứ không in năm: tên file mang đúng một token năm và đó là năm PHÁT HÀNH,
        # không phải năm dữ liệu ('...2026.7.18. BaocaoClaim_B2C_T11.25' là tháng 11 NĂM 2025).
        kys = sorted({k for _, k in ds if k != "?"})
        msg = (f"{muc}: {len(ds)} file chưa vào DB"
               + (f" (kỳ {', '.join(kys)})" if kys else "")
               + f" — vd {ds[0][0][:46]}")
        canh_bao.append(msg)
        ctx.log(f"  MỒ CÔI {msg}")
    return canh_bao


def _lat_du_lieu(ctx: Ctx, ctys: list) -> list:
    """[(rt, cong_ty, khoi, khoa, kieu_khoa, source_file, so_dong)] — một lát = một đơn vị đo bị
    cộng chung.

    KHOÁ LÁT KHÁC NHAU THEO BẢN CHẤT BÁO CÁO, không thể dùng chung một kiểu:
    · ẢNH CHỤP (`_SNAP_RT`) -> khoá là NGÀY CHỐT. Cả tháng có 4-5 bản tuần là bình thường (KSCL
      của XDV), chỉ hai file CÙNG một ngày chốt mới là cộng đôi.
    · BÁO CÁO DÒNG (claim, xuất hoá đơn…) -> khoá là KỲ. Cả tháng cộng dồn, nên hai bản chốt khác
      ngày của cùng một kỳ là cộng đôi — mà xét theo ngày thì chúng nằm hai ô khác nhau, lọt lưới.
      Đây KHÔNG phải giả định: prod 25/08 cõng claim T7 hai bản (13/08 + 25/08) cho cả B2B lẫn
      B2C, tức số claim tháng 7 đang gấp đôi, mà bản dò chỉ soi ảnh chụp thì không thấy gì.

    `_SNAP_RT` nhập THẲNG từ API (`app.metrics._shared`) chứ không chép sang đây: danh sách đó dài
    38 mã và còn dài ra mỗi lần thêm nguồn số dư, chép là cầm chắc có ngày lệch.

    Dòng của file đang ẨN (`hidden_files`) bị loại: ẩn nghĩa là mọi truy vấn của dashboard đã bỏ
    chúng ra: báo tiếp là báo một cái đã xử lý xong.
    """
    # SQL cố ý KHÔNG có COALESCE/chuỗi rỗng: mỗi dấu nháy trong đoạn code nhúng phải escape hai
    # tầng (chuỗi Python ở đây -> chuỗi Python bên trong `python -c`), sai một tầng là SyntaxError
    # nuốt cả bộ dò. Ghép giá trị rỗng bằng `or ''` ở tầng Python cho khỏi phải escape gì.
    code = (
        "import sys;sys.path.insert(0,'.');"
        "from app.metrics._shared import _SNAP_RT;"
        "from app.database.session import get_db;"
        f"ctys={sorted(ctys)!r};"
        "pc=','.join(['?']*len(ctys));"
        "sql=('SELECT report_type rt,cong_ty ct,khoi k,ngay ng,period_month pm,source_file sf,"
        "COUNT(*) c FROM raw_rows WHERE split_part(source_file,%s,1) IN (%s) AND source_file NOT "
        "IN (SELECT source_file FROM hidden_files) GROUP BY 1,2,3,4,5,6')%(chr(39)+'::'+chr(39),"
        "pc);"
        "\nfor r in get_db().execute(sql, tuple(ctys)):\n"
        "    snap=r['rt'] in _SNAP_RT\n"
        "    print('LAT|%s|%s|%s|%s|%s|%s|%s'%(r['rt'],r['ct'] or '',r['k'] or '',"
        "(r['ng'] if snap else r['pm']) or '','chốt' if snap else 'kỳ',r['sf'],r['c']))\n"
    )
    out = _py_sql(ctx, code, timeout=300)
    lat = []
    for line in out.splitlines():
        if line.startswith("LAT|"):
            p = line.split("|")
            if len(p) == 8 and p[4]:
                lat.append((p[1], p[2], p[3], p[4], p[5], p[6], int(p[7])))
        elif line.strip():
            ctx.log(f"  RÀ SOÁT CỘNG ĐÔI lỗi: {line[:200]}")
    return lat


def ra_soat_cong_doi(ctx: Ctx, nguon_list: list) -> list:
    """CỘNG ĐÔI = trong CÙNG một lát (xem `_lat_du_lieu`) có nhiều hơn một file cùng đóng góp.

    Vì sao không tự khỏi: với ảnh chụp, `_rows()` chọn ngày chốt theo (công ty, khối) chứ KHÔNG
    theo source_file nên hai file cùng ngày được cộng cả hai; với báo cáo dòng thì cộng dồn cả kỳ
    là đúng nghiệp vụ, hai bản chốt của cùng kỳ cũng cộng cả hai. `xoa_ban_cu` chỉ dọn được bản cũ
    mà lượt chạy NHÌN THẤY ở nguồn — bản mà kế toán đã xoá/đổi tên bên nguồn thì nằm lại trong DB
    vĩnh viễn, không lượt nào chọn nó làm "bản thua" nữa.

    Hai kiểu bị bắt, cùng là "một thứ đếm hai lần" nhưng cách sửa KHÁC hẳn nhau:
    · TRÙNG BẢN CHỐT — cùng tên gốc, khác ngày phát hành. Xoá rows của bản cũ (`source_file` nào
      không còn ở nguồn thì cron không bao giờ tự dọn được).
    · BẢN GỘP ⊃ BẢN TÁCH — file gộp cả kênh nằm cạnh các file tách kênh của đúng lát đó. Sửa bằng
      `bo_qua` cho tên bản gộp rồi xoá rows của nó; TUYỆT ĐỐI không gộp chung `slot` với các file
      kênh — chúng không phải hai bản chốt của nhau, gộp là xoá oan cả 3 kênh.
    """
    lat = _lat_du_lieu(ctx, [n["company"] for n in nguon_list])
    nhom = {}
    for rt, ct, khoi, khoa, kieu, sf, c in lat:
        nhom.setdefault((rt, ct, khoi, khoa, kieu), {}).setdefault(sf, 0)
        nhom[(rt, ct, khoi, khoa, kieu)][sf] += c
    canh_bao = []
    for (rt, ct, khoi, khoa, kieu), files in sorted(nhom.items()):
        if len(files) < 2:
            continue
        o = f"[{rt} {ct}/{khoi} {kieu} {khoa}]"
        theo_ten = {}
        for sf, c in files.items():
            theo_ten.setdefault(_ten_goc(sf), []).append((sf.split("::")[-1], c))
        for _, same in sorted(theo_ten.items()):
            if len(same) > 1:
                chi_tiet = " + ".join(f"{f} ({c} dòng)" for f, c in sorted(same))
                canh_bao.append(f"TRÙNG BẢN CHỐT {o}: {chi_tiet}")
        for a in sorted(theo_ten):
            con = [b for b in sorted(theo_ten) if b != a and _la_ban_gop(a, b)]
            if con:
                fa, ca = theo_ten[a][0]
                chi_tiet = ", ".join(f"{theo_ten[b][0][0]} ({theo_ten[b][0][1]} dòng)"
                                     for b in con)
                canh_bao.append(f"BẢN GỘP ⊃ BẢN TÁCH {o}: {fa} ({ca} dòng) nằm cạnh {chi_tiet}")
    for msg in canh_bao:
        ctx.log(f"  CỘNG ĐÔI {msg[:250]}")
    if not canh_bao:
        ctx.log(f"RÀ SOÁT CỘNG ĐÔI: {len(nhom)} lát dữ liệu, không lát nào bị 2 file cùng đóng góp")
    else:
        ctx.log(f"RÀ SOÁT CỘNG ĐÔI: {len(canh_bao)} lát bị đếm hai lần trên {len(nhom)} lát —"
                " SỐ TRÊN DASHBOARD ĐANG SAI, cách sửa xem docstring ra_soat_cong_doi()")
    return canh_bao


def ra_soat(ctx: Ctx, nguon_list: list, st=None) -> dict:
    """Chạy cả 2 bộ rà soát và ghi vào artifact trạng thái (agent gửi tin lãnh đạo đọc file này)."""
    kq = {"mo_coi": ra_soat_mo_coi(ctx, nguon_list),
          "cong_doi": ra_soat_cong_doi(ctx, nguon_list)}
    if st:
        st.set_run(ra_soat=kq)
    return kq

# ── vòng chạy ───────────────────────────────────────────────────────────────────────────────
def run(job: str, nhan: str, nguon_list: list, schedule_vn: str, argv=None) -> int:
    """Vòng chạy đầy đủ của một job. `nhan` = tên tiếng Việt cho log/tin gửi lãnh đạo."""
    import argparse
    for n in nguon_list:                     # gõ sai chế độ thì nổ NGAY, không im lặng bỏ nguồn
        if n["che_do"] not in CHE_DO:
            raise ValueError(f"che_do lạ {n['che_do']!r} ở nguồn {n['company']}/{n['rt']}")

    ap = argparse.ArgumentParser(description=f"Kéo + nạp nguồn {nhan}")
    ap.add_argument("--dry-run", action="store_true",
                    help="chỉ in bản sẽ chọn / bản sẽ loại, KHÔNG xin file, KHÔNG nạp, KHÔNG xoá")
    ap.add_argument("--env", choices=("test", "prod"), default="test")
    ap.add_argument("--chi-ra-soat", action="store_true",
                    help="CHỈ rà soát DB (file mồ côi + lát bị cộng đôi) rồi thoát — không xin"
                         " file, không nạp, không xoá, không ghi artifact trạng thái")
    args = ap.parse_args(argv)

    cfg = moi_truong(job)[args.env]
    ctx = Ctx(job, args.env, cfg, args.dry_run)
    ctx.log("=" * 70)
    ctx.log(f"{nhan} — MOI TRUONG: {args.env} (DB {cfg['database_url'].rsplit('@', 1)[-1]})")

    if args.chi_ra_soat:
        # Cố ý KHÔNG dựng StatusWriter ở nhánh này: ghi artifact là báo với agent gửi tin lãnh đạo
        # rằng "job vừa chạy", trong khi đây chỉ là một lượt soi DB do người gõ tay.
        ra_soat(ctx, nguon_list)
        return 0

    # Mẫu số của artifact = DANH SÁCH KHAI BÁO, không phải số file tìm thấy: thư mục vắng mặt hoàn
    # toàn khỏi metadata không để lại dấu vết nào trong log (bài học DUAN 22/08/2026).
    expected = [f"{n['company']}/{n['rt']}" for n in nguon_list]
    ten = {f"{n['company']}/{n['rt']}": n.get("ten") or f"{n['company']}/{n['rt']}"
           for n in nguon_list}
    st = None if args.dry_run else cron_status.StatusWriter(
        job=job, env=args.env, json_path=cfg["status_json"], jsonl_path=cfg["status_jsonl"],
        schedule_vn=schedule_vn, expected=sorted(expected), names=ten)

    def done(status=cron_status.RUN_OK, note=None, rc=0):
        if st:
            st.finish(status, note)
        return rc

    if os.path.exists(cfg["disabled_flag"]) and not args.dry_run:
        ctx.log("BO QUA: dang TAT tu giao dien (Nguon du lieu > Keo tu dong). Xoa co de bat lai.")
        return done(cron_status.RUN_TAT, "job dang bi TAT tu giao dien Nguon du lieu")

    periods = target_periods()
    ctx.log(f"BẮT ĐẦU — {len(nguon_list)} thư mục nguồn, kỳ cần kéo: {[p for p, _, _ in periods]}"
            + (" [DRY-RUN]" if args.dry_run else ""))
    if st:
        st.set_run(periods=[p for p, _, _ in periods], period_chinh=periods[0][0])

    before_meta = meta_mtime()
    try:
        http("POST", "/request-metadata")
    except (urllib.error.URLError, OSError) as ex:
        ctx.log(f"DỪNG: không gọi được receiver {RECEIVER_URL} ({ex}). receiver_server còn chạy?")
        return done(cron_status.RUN_DUNG_SOM, f"không gọi được receiver: {str(ex)[:120]}", 1)
    if args.dry_run:
        ctx.log("DRY-RUN: không chờ, đọc luôn danh sách đang có (có thể là bản quét cũ)")
    else:
        wait_metadata(ctx, before_meta)

    meta = read_metadata(ctx)
    if meta is None:
        return done(cron_status.RUN_DUNG_SOM, "không đọc được available_metadata.json", 1)

    targets, losers = pick_targets(ctx, nguon_list, meta, periods)
    for t in targets:
        t["_slot_key"] = _slot(t["_nguon"], t, t["_period"])
    for o in losers:
        o["_slot_key"] = _slot(o["_nguon"], o, o["_period"])

    ky_chinh = periods[0][0]
    thay = {f"{t.get('company')}/{t.get('report_type')}" for t in targets}
    # PHÂN BIỆT "KHÔNG CÓ FILE NÀO" VỚI "CHƯA CÓ FILE KỲ NÀY". Bản đầu ghi chung một câu "chưa
    # thấy file báo cáo ở nguồn" cho cả hai, nên 4 mục (claim B2B, công nợ phải thu, nhập xe
    # B2B/B2C) hiện y như thể nguồn trống trơn — trong khi chúng CÓ file, chỉ là mới tới T07.
    # Câu này đi thẳng vào tin gửi lãnh đạo nên sai nghĩa là họ đi hỏi kế toán sai chỗ.
    thay_ky_chinh = {f"{t.get('company')}/{t.get('report_type')}"
                     for t in targets if t["_period"] == ky_chinh}
    for muc in sorted(set(expected) - thay_ky_chinh):
        ky_khac = sorted({t["_period"] for t in targets
                          if f"{t.get('company')}/{t.get('report_type')}" == muc})
        if muc in thay:
            ctx.log(f"  CHUA CO FILE KY {ky_chinh}: {muc} (nguồn mới có kỳ {', '.join(ky_khac)})")
        else:
            ctx.log(f"  KHONG THAY FILE NAO: {muc} (kỳ {[p for p, _, _ in periods]})")
        if st:
            st.record(muc, state=cron_status.STATE_CHUA_CO_FILE, arrived=False,
                      ly_do=(f"kỳ {ky_chinh} chưa có file ở nguồn (mới có kỳ"
                             f" {', '.join(ky_khac)})" if muc in thay
                             else "chưa thấy file báo cáo nào ở nguồn"))
    if not targets:
        ctx.log("DỪNG: không có file nào khớp kỳ cần kéo")
        return done(cron_status.RUN_DUNG_SOM, "không có file nào khớp kỳ cần kéo", 1)

    # Bỏ ảnh chụp tuần CŨ đã có dòng trong DB (xem GIU_MOI_NHAT). Làm TRƯỚC khi xin file để không
    # tốn cả lượt truyền file lẫn lượt derive.
    co_the_bo = [e for e in targets if e.get("_bo_qua_neu_co_roi")]
    if co_the_bo and not args.dry_run:
        da_co = da_co_dong(ctx, [source_id(e) for e in co_the_bo])
        bo = [e for e in co_the_bo if source_id(e) in da_co]
        for e in bo:
            ctx.log(f"  bỏ qua (đã nạp, không phải bản mới nhất): {e['fileName'][:50]}")
        targets = [e for e in targets if e not in bo]

    for e in targets:
        ctx.log(f"  chọn [{e['_period']}] [{e.get('report_type')}] {e['fileName'][:48]}"
                f" modifiedAt={e.get('modifiedAt')}")
    if args.dry_run:
        ctx.log(f"DRY-RUN: sẽ nạp {len(targets)} file (trong đó {len(co_the_bo)} bản ảnh chụp cũ sẽ"
                f" bỏ nếu DB đã có), sẽ xoá rows của {len(losers)} bản cũ. Dừng ở đây.")
        return 0

    ctx.log(f"KÉO {len(targets)} file ({len({e.get('report_type') for e in targets})} thư mục)")
    before = request_files(ctx, targets)
    arrived = wait_arrival(ctx, targets, before)

    today = datetime.now(VN).strftime("%Y-%m-%d")
    ok, da_nap_ok = 0, set()
    for e in sorted(targets, key=lambda x: (x["_period"], x.get("report_type") or "")):
        muc = f"{e.get('company')}/{e.get('report_type')}"
        rec = st.record if (st and e["_period"] == ky_chinh) else (lambda *a, **k: None)
        if not os.path.exists(xlsx_path(e)):
            ctx.log(f"  BỎ QUA [{muc}] {e['_period']}: không có file trên đĩa")
            rec(muc, state=cron_status.STATE_CHUA_CO_FILE, arrived=False,
                ly_do="đã xin file nhưng không về, trên đĩa cũng chưa có bản nào")
            continue
        af_ok, af_err, af_cb = autofill(ctx, e)
        for c in af_cb:
            ctx.log(f"    canh bao: {str(c)[:200]}")
        if not af_ok:
            rec(muc, state=cron_status.STATE_LOI_NAP, arrived=e["fileName"] in arrived,
                ly_do="hệ thống không đọc được số liệu trong file, cần kế toán kiểm tra lại file",
                ly_do_ky_thuat=af_err)
            continue
        vr = verify(ctx, e)
        if vr.get("code") == "KHONG_CO_DONG_NAO":
            rec(muc, state=cron_status.STATE_LOI_NAP, arrived=e["fileName"] in arrived,
                ly_do="có file nhưng nạp ra 0 dòng", ly_do_ky_thuat=vr.get("raw"))
            continue
        ok += 1
        da_nap_ok.add(e["_slot_key"])
        state, doi_luc = trang_thai(vr, st.van_tay_cu.get(muc) if st else None,
                                    st.doi_luc_cu.get(muc) if st else None, today)
        mx = vr.get("max_ngay")
        rec(muc, state=state, arrived=e["fileName"] in arrived, so_ngay=vr.get("so_ngay"),
            max_ngay=mx, verify_code=vr.get("code"), van_tay=vr.get("van_tay"), doi_luc=doi_luc,
            ly_do=(f"đã có số, bản chốt {mx[8:10]}/{mx[5:7]}" if state == cron_status.STATE_DU and mx
                   else "đã có số" if state == cron_status.STATE_DU
                   else "file chưa đổi kể từ lượt trước, kế toán chưa cập nhật"
                        + (f" (đổi gần nhất {doi_luc[8:10]}/{doi_luc[5:7]})" if doi_luc else "")))

    so_xoa = xoa_ban_cu(ctx, losers, da_nap_ok) if losers else 0
    # SAU khi đã nạp + đã xoá bản chốt cũ: rà lại chính cái vừa để lại trong DB. Đặt ở đây chứ
    # không phải đầu lượt để không báo động cái mà `xoa_ban_cu` của chính lượt này vừa dọn xong.
    rs = ra_soat(ctx, nguon_list, st)
    ctx.log(f"XONG — nạp thành công {ok}/{len(targets)} file"
            + (f", xoá {so_xoa} dòng của {len(losers)} bản chốt cũ" if losers else "")
            + (f", RÀ SOÁT: {len(rs['cong_doi'])} lát cộng đôi / {len(rs['mo_coi'])} nhóm file mồ"
               " côi" if (rs["cong_doi"] or rs["mo_coi"]) else ", rà soát sạch"))
    if st:
        st.set_run(nap_thanh_cong=ok, so_file_keo=len(targets))
    return done(cron_status.RUN_OK, rc=0 if ok else 1)

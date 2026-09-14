# -*- coding: utf-8 -*-
"""KÉO FILE NGUỒN từ máy Local qua receiver (POST /request-file) — helper dùng chung cho các cron
Dòng tiền (số dư ngân hàng, vay theo ngày).

VÌ SAO KHÔNG import cron_thuchi_daily: module đó mang global state của riêng job nó (LOG_FILE,
DATABASE_URL... gán lại trong main()), import vào là kéo theo cả đống phụ thuộc + ghi nhầm log.

VÌ SAO KHÔNG DÙNG `sync_orchestrator.py pull`: cmd_plan chỉ xin file CHƯA nhận, khoá theo
(company, fileName) — mà các file này GIỮ NGUYÊN TÊN và bị GHI ĐÈ mỗi ngày, nên sau lần đầu nó
bị coi là "đã nhận" vĩnh viễn. Ở đây POST /request-file thẳng từ available_metadata.json để ép kéo
lại, giống hệt cách cron_thuchi_daily.py đang làm (receiver ghi bằng "wb" nên đè an toàn).

KHÔNG refresh metadata mặc định: danh sách file ở đây CỐ ĐỊNH (12 file số dư, các file báo cáo
ngân hàng theo tháng) — chỉ NỘI DUNG đổi, không phát sinh tên mới. Refresh tốn ~70s quét bên máy
Local mà không đem lại gì. Bật `refresh=True` nếu có nguồn mới chưa thấy trong danh sách.
"""
import json
import os
import time
import urllib.error
import urllib.request

ROOT = "/home/itadmin/AI_Dashboard_QT"
CONNECT = f"{ROOT}/Connect_VPS"
AVAILABLE_META = f"{CONNECT}/available_metadata.json"
RECEIVED_DIR = f"{CONNECT}/received_reports"
RECEIVER_URL = "http://127.0.0.1:8090"

ARRIVE_TIMEOUT = 300        # giây, chờ file thực sự về đĩa (cùng mốc cron_thuchi_daily)
ARRIVE_POLL = 5
META_REFRESH_TIMEOUT = 240
META_REFRESH_POLL = 5

_PAYLOAD_KEYS = ("company", "report_type", "fileName", "path", "month", "periodType", "status")


def _http(method: str, path: str, payload=None, timeout=20):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(RECEIVER_URL + path, data=data, method=method,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        body = r.read().decode(errors="replace")
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return {"raw": body[:200]}


def read_metadata():
    """available_metadata.json -> list entry. Thử lại nếu trúng lúc receiver đang ghi dở."""
    for _ in range(3):
        try:
            with open(AVAILABLE_META, encoding="utf-8") as fh:
                m = json.load(fh)
            return m if isinstance(m, list) else (m.get("files") or m.get("items") or [])
        except json.JSONDecodeError:
            time.sleep(2)
        except OSError:
            return None
    return None


def refresh_metadata(log=print) -> bool:
    """Bảo máy Local quét lại rồi CHỜ tới khi available_metadata.json thực sự đổi mtime."""
    try:
        truoc = os.path.getmtime(AVAILABLE_META)
    except OSError:
        truoc = None
    try:
        _http("POST", "/request-metadata")
    except (urllib.error.URLError, OSError) as ex:
        log(f"  LỖI xin danh sách: {ex}")
        return False
    het = time.time() + META_REFRESH_TIMEOUT
    while time.time() < het:
        time.sleep(META_REFRESH_POLL)
        try:
            if os.path.getmtime(AVAILABLE_META) != truoc:
                return True
        except OSError:
            pass
    log(f"  danh sách KHÔNG đổi sau {META_REFRESH_TIMEOUT}s — dùng bản hiện có")
    return False


def _sidecar_saved_at(e: dict):
    """saved_at trong sidecar .json cạnh file đã nhận; None = chưa từng nhận."""
    p = os.path.join(RECEIVED_DIR, e.get("company") or "", e.get("report_type") or "",
                     os.path.splitext(e["fileName"])[0] + ".json")
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh).get("saved_at")
    except (OSError, json.JSONDecodeError):
        return None


def keo(chon, log=print, refresh: bool = False, timeout: int = ARRIVE_TIMEOUT) -> dict:
    """Kéo mọi file mà `chon(entry)` trả True. -> {'xin': n, 've': n, 'thieu': [fileName...]}.

    CHỜ file thực sự về (sidecar saved_at đổi) mới trả: /request-file chỉ XẾP HÀNG, máy Local
    poll 2s rồi mới upload — không chờ thì bước trích xuất ngay sau đó xử lý BẢN CŨ mà vẫn báo OK.
    File không về đúng hạn KHÔNG làm hỏng lượt chạy: bước sau vẫn đọc bản cũ trên đĩa (và log đã
    nói rõ là bản cũ), đúng tinh thần "thà số cũ còn hơn màn trắng".
    """
    if refresh:
        refresh_metadata(log)
    meta = read_metadata()
    if meta is None:
        log(f"  DỪNG KÉO: không đọc được {AVAILABLE_META}")
        return {"xin": 0, "ve": 0, "thieu": []}
    targets = [e for e in meta if e.get("fileName") and chon(e)]
    if not targets:
        log("  KHÔNG có file nào khớp điều kiện kéo")
        return {"xin": 0, "ve": 0, "thieu": []}

    truoc = {}
    for e in targets:
        truoc[e["fileName"]] = _sidecar_saved_at(e)
        try:
            _http("POST", "/request-file", {k: e.get(k) for k in _PAYLOAD_KEYS})
        except (urllib.error.URLError, OSError) as ex:
            log(f"  LỖI xin {e['fileName'][:50]}: {ex}")
    log(f"  đã xin {len(targets)} file, chờ về (tối đa {timeout}s)…")

    con_lai = {e["fileName"]: e for e in targets}
    het = time.time() + timeout
    ve = 0
    while con_lai and time.time() < het:
        time.sleep(ARRIVE_POLL)
        for fn, e in list(con_lai.items()):
            moi = _sidecar_saved_at(e)
            if moi and moi != truoc.get(fn):
                ve += 1
                del con_lai[fn]
                log(f"  đã về: {fn[:60]} ({moi})")
    for fn in con_lai:
        log(f"  KHÔNG VỀ sau {timeout}s: {fn[:55]}"
            + (" — dùng bản cũ trên đĩa" if truoc.get(fn) else " — chưa từng có bản nào"))
    return {"xin": len(targets), "ve": ve, "thieu": sorted(con_lai)}

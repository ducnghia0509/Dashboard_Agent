# -*- coding: utf-8 -*-
"""Daemon giữ BRIEF_CHUTICH.xlsx LUÔN NÓNG — yêu cầu của Chủ tịch: agent phải ở trạng thái
đã cập nhật số liệu, chờ sẵn để trả lời, không phải đi tính lúc được hỏi.

Cơ chế: `Connect_VPS/receiver_server.py` (đang chạy) mỗi lần nhận file đều ghi 1 dòng vào
`inbox_events.jsonl`. Hook này hiện KHÔNG có ai tiêu thụ (docstring nói sync_orchestrator xử lý
nhưng không có code nào đọc) — dùng luôn làm nguồn trigger, không phải dựng thêm hạ tầng.

Ba lớp bảo vệ:
  1. Sự kiện: theo dõi inbox_events.jsonl, gom cụm (debounce) rồi dựng lại ĐÚNG sheet bị ảnh hưởng.
  2. Định kỳ: dựng full lúc khởi động + mỗi FULL_REBUILD_HOURS giờ — bắt file copy tay (không
     qua receiver nên không sinh event).
  3. Cờ STALE: build_chutich_brief tự ghi vào _META khi sheet cũ nhất quá 2 giờ; SKILL.md buộc
     agent nói ra thay vì trả số cũ như số mới.

Chạy tay:   .venv/bin/python scripts/chutich_brief_watcher.py --once
Chạy nền:   systemctl --user start chutich-brief   (xem deploy/chutich-brief.service)
"""
import argparse
import json
import os
import sys
import time
import traceback
from datetime import datetime, timedelta, timezone

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
for _p in (_ROOT, _HERE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from scripts import build_chutich_brief as B  # noqa: E402

INBOX_EVENTS = B.INBOX_EVENTS
STATE_PATH = os.path.join(_ROOT, "memory", "chutich_brief_watcher.state.json")

DEBOUNCE_SECONDS = 90          # nhịp thật: 3 file về trong 2 giây -> gom lại 1 lần dựng
POLL_SECONDS = 5
FULL_REBUILD_HOURS = 1
TZ = B.TZ


def _log(msg):
    print(f"[{datetime.now(TZ).strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


def _load_state():
    try:
        with open(STATE_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"offset": 0, "inode": None, "last_full": None}


def _save_state(st):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    tmp = STATE_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE_PATH)


def _read_new_events(st):
    """Đọc phần MỚI của inbox_events.jsonl từ byte-offset đã lưu.

    Bám theo inode: file bị xoay vòng/thay mới thì inode đổi -> đọc lại từ đầu thay vì giữ
    offset cũ (giữ offset của file cũ sẽ nhảy quá phần đầu file mới, mất event im lặng).
    """
    try:
        stat = os.stat(INBOX_EVENTS)
    except FileNotFoundError:
        return [], st
    inode = stat.st_ino
    offset = st.get("offset", 0)
    if st.get("inode") != inode or offset > stat.st_size:
        _log(f"inbox_events đổi/ngắn lại (inode {st.get('inode')} -> {inode}) — đọc lại từ đầu")
        offset = 0
    if offset == stat.st_size:
        return [], {**st, "inode": inode, "offset": offset}
    with open(INBOX_EVENTS, "rb") as fh:
        fh.seek(offset)
        chunk = fh.read()
        new_offset = fh.tell()
    # Dòng cuối có thể đang được ghi dở -> chỉ xử lý tới ký tự xuống dòng cuối cùng.
    nl = chunk.rfind(b"\n")
    if nl == -1:
        return [], {**st, "inode": inode, "offset": offset}
    new_offset = offset + nl + 1
    events = []
    for line in chunk[:nl].decode("utf-8", "ignore").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            events.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return events, {**st, "inode": inode, "offset": new_offset}


def _sheets_for(events):
    """report_type của các file vừa về -> tập sheet cần dựng lại."""
    sheets, unknown = set(), set()
    for ev in events:
        rt = ev.get("report_type")
        target = B.REPORT_TYPE_SHEETS.get(rt)
        if target:
            sheets.update(target)
        elif rt:
            unknown.add(rt)
    return sheets, unknown


def _rebuild(sheets=None, why=""):
    t0 = time.time()
    label = ", ".join(sorted(sheets)) if sheets else "TOÀN BỘ"
    _log(f"dựng lại [{label}] — {why}")
    try:
        r = B.build(only=sorted(sheets) if sheets else None, verbose=False)
    except Exception:
        _log("LỖI khi dựng brief:\n" + traceback.format_exc())
        return False
    _log(f"xong sau {time.time() - t0:.1f}s · sheet={len(r['sheets'])} "
         f"· thiếu nguồn={len(r['notes'])} · lỗi={len(r['errors'])}"
         + (f" · {r['stale']}" if r["stale"] else ""))
    for n in r["notes"]:
        _log(f"  [THIẾU] {n}")
    for e in r["errors"]:
        _log(f"  [LỖI] {e}")
    return True


def run(once=False, catch_up=False):
    st = _load_state()
    if not catch_up and st.get("offset", 0) == 0 and os.path.exists(INBOX_EVENTS):
        # Lần đầu chạy: bỏ qua toàn bộ lịch sử event (838+ dòng) — dựng full là đủ, không cần
        # phát lại từng file cũ.
        st["offset"] = os.path.getsize(INBOX_EVENTS)
        st["inode"] = os.stat(INBOX_EVENTS).st_ino
        _log(f"lần đầu: bỏ qua {st['offset']} byte lịch sử, chỉ theo dõi event mới")

    _rebuild(None, "khởi động")
    st["last_full"] = datetime.now(TZ).isoformat()
    _save_state(st)
    if once:
        return 0

    pending, pending_since = set(), None
    while True:
        time.sleep(POLL_SECONDS)
        try:
            events, st2 = _read_new_events(st)
            if events:
                sheets, unknown = _sheets_for(events)
                files = ", ".join(e.get("file_name", "?") for e in events[:5])
                _log(f"{len(events)} file mới ({files}{'…' if len(events) > 5 else ''})")
                if unknown:
                    _log(f"  report_type chưa map sang sheet nào: {sorted(unknown)} "
                         f"— sẽ được cập nhật ở lần dựng full định kỳ")
                if sheets:
                    pending |= sheets
                    pending_since = pending_since or time.time()
                st = st2
                _save_state(st)

            if pending and pending_since and time.time() - pending_since >= DEBOUNCE_SECONDS:
                _rebuild(pending, f"gom {DEBOUNCE_SECONDS}s sau khi có file mới")
                pending, pending_since = set(), None

            last_full = st.get("last_full")
            due = True
            if last_full:
                try:
                    due = datetime.now(TZ) - datetime.fromisoformat(last_full) >= timedelta(
                        hours=FULL_REBUILD_HOURS)
                except ValueError:
                    due = True
            if due:
                _rebuild(None, f"định kỳ mỗi {FULL_REBUILD_HOURS}h (bắt file copy tay)")
                st["last_full"] = datetime.now(TZ).isoformat()
                _save_state(st)
        except KeyboardInterrupt:
            _log("dừng theo yêu cầu")
            return 0
        except Exception:
            _log("LỖI vòng lặp (bỏ qua, chạy tiếp):\n" + traceback.format_exc())


def main():
    ap = argparse.ArgumentParser(description="Giữ BRIEF_CHUTICH.xlsx luôn cập nhật")
    ap.add_argument("--once", action="store_true", help="dựng full 1 lần rồi thoát")
    ap.add_argument("--catch-up", action="store_true",
                    help="lần đầu chạy vẫn đọc lại toàn bộ lịch sử event thay vì bỏ qua")
    a = ap.parse_args()
    return run(once=a.once, catch_up=a.catch_up)


if __name__ == "__main__":
    sys.exit(main())

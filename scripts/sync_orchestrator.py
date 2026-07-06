# -*- coding: utf-8 -*-
"""SYNC ORCHESTRATOR (Phase 3) — điều phối kéo file mới từ Connect_VPS về + index + xử lý.

Mô hình pull-via-push của Connect_VPS: server này xếp task, máy Local poll rồi upload file.
Orchestrator (chạy theo timer systemd hoặc gọi tay) làm:
  refresh  -> POST /request-metadata (bảo Local quét & gửi danh sách file mới nhất)
  plan     -> diff available_metadata vs file đã nhận + catalog -> {to_request, to_process}
  pull     -> POST /request-file cho từng file CHƯA nhận (Local sẽ upload)
  index    -> quét received_reports -> cập nhật source_catalog (QA tra được ngay)
  process  -> file đã nhận + có report_spec ĐÃ HỌC -> điền template + import; còn lại -> flag analyst
  run      -> index + plan (mặc định, an toàn; không tự request)

KHÔNG deps ngoài stdlib (urllib) cho HTTP. Chạy: .venv/bin/python scripts/sync_orchestrator.py <cmd>
"""
import argparse
import json
import os
import sys
import urllib.request

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from dotenv import load_dotenv  # noqa: E402
load_dotenv(os.path.join(_ROOT, ".env"))

from servers.common import source_catalog as SC  # noqa: E402

RECEIVER_URL = os.environ.get("RECEIVER_URL", "http://127.0.0.1:8090")
CONNECT_ROOT = os.path.normpath(os.path.join(_ROOT, "..", "Connect_VPS"))
AVAILABLE_META = os.environ.get("RECEIVER_META") or os.path.join(CONNECT_ROOT, "available_metadata.json")
SYNC_LEDGER = os.path.join(_ROOT, "memory", "sync_ledger.json")


RECEIVER_TOKEN = os.environ.get("RECEIVER_TOKEN", "").strip()


def _http(method: str, path: str, payload: dict = None, timeout: int = 15):
    url = RECEIVER_URL.rstrip("/") + path
    data = json.dumps(payload).encode() if payload is not None else None
    headers = {"Content-Type": "application/json"}
    if RECEIVER_TOKEN:
        headers["X-Auth-Token"] = RECEIVER_TOKEN
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        body = r.read().decode()
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return {"raw": body}


def _load_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return default


def _guess_period(file_name: str):
    """Suy kỳ 'YYYY-MM' từ tên file: 'YYYYMM'/'YYYY-MM' (202601) hoặc 'MM.YYYY' (05.2026)."""
    import re
    if not file_name:
        return None
    m = re.search(r"(20\d{2})[.\-_]?(0[1-9]|1[0-2])", file_name)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"(0[1-9]|1[0-2])[.\-_](20\d{2})", file_name)
    if m:
        return f"{m.group(2)}-{m.group(1)}"
    return None


def _period_for(file_name: str, avail: dict):
    """Kỳ 'YYYY-MM': ưu tiên suy từ tên file; nếu tên không có kỳ (vd 'M000526') -> dùng
    'month' trong metadata + năm (quét 20xx trong tên, mặc định env DASHBOARD_YEAR/2026)."""
    p = _guess_period(file_name)
    if p:
        return p
    import re
    m = (avail.get(file_name, {}) or {}).get("month")
    if not m:
        return None
    ym = re.search(r"20\d{2}", file_name or "")
    year = ym.group(0) if ym else os.environ.get("DASHBOARD_YEAR", "2026")
    return f"{year}-{int(m):02d}"


def _received_basenames() -> set:
    root = SC.RECEIVED_DIR
    out = set()
    if os.path.isdir(root):
        for dp, _, fs in os.walk(root):
            for f in fs:
                if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    out.add(f)
    return out


def cmd_refresh(_):
    return _http("POST", "/request-metadata")


def cmd_index(_):
    return SC.index_dir()


def cmd_plan(_):
    avail = _load_json(AVAILABLE_META, [])
    recv = _received_basenames()
    to_request = [e for e in avail if e.get("fileName") and e["fileName"] not in recv]
    to_process = [{"file": c["file"], "company": c.get("company"),
                   "canonical_kinds": sorted({s.get("canonical_kind") for s in c.get("sheets", [])
                                              if s.get("canonical_kind")})}
                  for c in SC.search(only_uningested=True)]
    return {"available": len(avail), "received": len(recv),
            "to_request": to_request, "to_process": to_process}


def cmd_pull(args):
    """Request file chưa nhận (Local upload dần). --files fn1 fn2 = CHỈ kéo các file này;
    không có --files: limit<=0 = kéo tất cả, limit>0 = tối đa limit."""
    plan = cmd_plan(args)
    wanted = set(getattr(args, "files", None) or [])
    if wanted:
        reqs = [e for e in plan["to_request"] if e.get("fileName") in wanted]
    else:
        lim = getattr(args, "limit", 0) or 0
        reqs = plan["to_request"] if lim <= 0 else plan["to_request"][:lim]
    results = []
    for e in reqs:
        payload = {"company": e.get("company"), "report_type": e.get("report_type"),
                   "fileName": e.get("fileName"), "path": e.get("path"),
                   "month": e.get("month"), "periodType": e.get("periodType", ""),
                   "status": e.get("status", "")}
        try:
            results.append({"fileName": e.get("fileName"), "resp": _http("POST", "/request-file", payload)})
        except Exception as ex:
            results.append({"fileName": e.get("fileName"), "error": str(ex)})
    return {"requested": len(results), "results": results}


def cmd_process(args):
    """TỰ fill+import các file đã nhận có mapping ĐÃ HỌC (fill_spec). Layout mới -> bỏ qua (analyst)."""
    from servers import template_filler as TF
    results = []
    for e in SC.search(only_uningested=True):
        path = e.get("path")
        if not path or not os.path.exists(path):
            continue
        try:
            r = TF.autofill_file(path, cong_ty=e.get("company"))
        except Exception as ex:
            results.append({"file": e["file"], "error": str(ex)})
            continue
        if r.get("any_processed"):
            SC.mark_ingested(path, True)
        results.append({"file": e["file"], "processed": len(r["processed"]),
                        "skipped_sheets": len(r["skipped_sheets"])})
    return {"processed_files": results}


def cmd_reprocess(args):
    """NẠP LẠI mọi file đã nhận có mapping ĐÃ HỌC — KHÔNG bỏ qua file đã ingested (khác cmd_process).
    Dùng khi đã XOÁ dữ liệu DB và cần dựng lại dashboard từ các file đã phân tích trước đó.
    --files fn1 fn2 = chỉ nạp lại các file này (không có -> tất cả)."""
    from servers import template_filler as TF
    avail = {a.get("fileName"): a for a in _load_json(AVAILABLE_META, []) if a.get("fileName")}
    only = set(getattr(args, "files", None) or [])
    results = []
    for e in SC.search():                       # KHÔNG lọc uningested -> gồm cả file "đã hiển thị"
        fn = e.get("file")
        path = e.get("path")
        if only and fn not in only:
            continue
        if not path or not os.path.exists(path):
            results.append({"file": fn, "error": "file không còn trên đĩa"})
            continue
        period = _period_for(fn, avail)
        try:
            r = TF.autofill_file(path, period=period, cong_ty=e.get("company"))
        except Exception as ex:
            results.append({"file": fn, "error": str(ex)})
            continue
        if r.get("any_processed"):
            SC.mark_ingested(path, True)
        results.append({"file": fn, "period": period, "any": bool(r.get("any_processed")),
                        "processed": len(r["processed"]), "skipped_sheets": len(r["skipped_sheets"])})
    return {"reprocessed": results,
            "ok_count": sum(1 for x in results if x.get("any")),
            "skipped_no_spec": sum(1 for x in results if x.get("any") is False and "error" not in x)}


def cmd_run(args):
    """An toàn: index catalog + plan (KHÔNG tự request/ghi). Dùng khi chỉ muốn khảo sát."""
    idx = SC.index_dir()
    plan = cmd_plan(args)
    return {"index": idx, "plan": {k: plan[k] for k in ("available", "received")},
            "to_request_count": len(plan["to_request"]),
            "to_process_count": len(plan["to_process"]),
            "to_process": plan["to_process"]}


def cmd_auto(args):
    """Timer: index -> process (fill+import file đã học) -> tóm tắt còn lại. Layout mới vẫn chờ analyst."""
    idx = SC.index_dir()
    proc = cmd_process(args)
    plan = cmd_plan(args)
    return {"index": idx, "processed": proc["processed_files"],
            "to_request_count": len(plan["to_request"]),
            "still_to_process": len(plan["to_process"])}


def cmd_mark(args):
    """Đánh dấu 1 file (theo path) đã ingested trong catalog — dùng sau khi analyst import xong."""
    SC.mark_ingested(args.path, True)
    return {"ok": True, "path": args.path, "ingested": True}


def cmd_status(args):
    """Danh sách nguồn cho UI: mỗi file + trạng thái (mới / chưa nạp / đã phân tích & hiển thị).

    - new       : có trong available_metadata nhưng CHƯA kéo về (received=False)
    - pending   : đã kéo về nhưng CHƯA import (received=True, ingested=False)
    - ingested  : đã phân tích & lên dashboard (ingested=True)
    """
    SC.index_dir()  # cập nhật catalog trước khi báo cáo
    avail = {a.get("fileName"): a for a in _load_json(AVAILABLE_META, []) if a.get("fileName")}
    cat = {e["file"]: e for e in SC.search()}
    names = set(avail) | set(cat)
    files = []
    for fn in sorted(names):
        a = avail.get(fn, {})
        c = cat.get(fn)
        received = c is not None
        ingested = bool(c and c.get("ingested"))
        state = "ingested" if ingested else ("pending" if received else "new")
        label = {"ingested": "Đã phân tích & hiển thị",
                 "pending": "Đã kéo về — chờ phân tích",
                 "new": "Mới — chưa kéo về"}[state]
        files.append({
            "file": fn,
            "company": (c or a).get("company"),
            "report_type": (c or a).get("report_type"),
            "month": a.get("month"),
            "period_type": a.get("periodType") or a.get("period_type"),
            "state": state, "label": label,
            "path": c.get("path") if c else None,          # đường dẫn file đã kéo về (để analyst đọc)
            "period": _guess_period(fn),                   # 'YYYY-MM' suy từ tên file
            "canonical_kinds": sorted({s.get("canonical_kind") for s in (c or {}).get("sheets", [])
                                       if s.get("canonical_kind")}) if c else [],
        })
    summary = {"total": len(files),
               "new": sum(1 for f in files if f["state"] == "new"),
               "pending": sum(1 for f in files if f["state"] == "pending"),
               "ingested": sum(1 for f in files if f["state"] == "ingested")}
    return {"summary": summary, "files": files}


def main():
    ap = argparse.ArgumentParser(description="Sync orchestrator Connect_VPS -> catalog/ingest")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("refresh", "index", "plan", "run", "process", "auto", "status"):
        sub.add_parser(name)
    p_pull = sub.add_parser("pull")
    p_pull.add_argument("--limit", type=int, default=0)  # 0 = kéo tất cả file mới
    p_pull.add_argument("--files", nargs="*", default=None)  # chỉ kéo các fileName này
    p_re = sub.add_parser("reprocess")
    p_re.add_argument("--files", nargs="*", default=None)  # chỉ nạp lại các fileName này
    p_mark = sub.add_parser("mark-ingested")
    p_mark.add_argument("path")
    args = ap.parse_args()
    fn = {"refresh": cmd_refresh, "index": cmd_index, "plan": cmd_plan,
          "pull": cmd_pull, "run": cmd_run, "process": cmd_process, "auto": cmd_auto,
          "status": cmd_status, "mark-ingested": cmd_mark, "reprocess": cmd_reprocess}[args.cmd]
    print(json.dumps(fn(args), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

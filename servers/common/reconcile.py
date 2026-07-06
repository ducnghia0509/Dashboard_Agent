# -*- coding: utf-8 -*-
"""RECONCILER (Phase 4) — soi lỗ hổng của pipeline hợp nhất, KHÔNG ghi gì:
- file đã KÉO VỀ nhưng CHƯA import (catalog.ingested=False) -> QA không mù, admin biết cần xử.
- KPI/màn FE THIẾU NGUỒN: report_type builder cần chưa có dữ liệu trong dataset.
- pipeline_state: trace từng file collected -> received(indexed) -> ingested (1 nguồn sự thật hợp nhất).
"""
import json
import os

from . import be_bridge as bb
from . import source_catalog as SC

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
CONNECT_ROOT = os.path.normpath(os.path.join(_AGENT_ROOT, "..", "Connect_VPS"))
AVAILABLE_META = os.environ.get("RECEIVER_META") or os.path.join(CONNECT_ROOT, "available_metadata.json")

# KPI -> report_type nguồn (khớp metrics.KPI_SOURCE của DashBoard_AI).
_KPI_SOURCE = {"doanh_thu": "DTHU", "lntt": "HQKD", "chi_phi": "HQKD", "dong_tien": "THUCHI",
               "tien_nh": "SDT", "cong_no": "PTHU", "ton_kho": "HH", "tai_san": "TS"}


def _active_id():
    for kind in ("month", "day", "generic"):
        d = bb.repo.active_dataset(kind)
        if d:
            return d["id"]
    return None


def uningested_files() -> list:
    """File đã index vào catalog nhưng chưa import raw_rows (còn chờ analyst/auto-process)."""
    return [{"file": e["file"], "company": e.get("company"),
             "canonical_kinds": sorted({s.get("canonical_kind") for s in e.get("sheets", []) if s.get("canonical_kind")})}
            for e in SC.search(only_uningested=True)]


def screens_missing_source(dataset_id: str = None) -> dict:
    """report_type builder cần nhưng dataset chưa có -> KPI/màn 'chưa có dữ liệu'."""
    ds = dataset_id or _active_id()
    if not ds:
        return {"dataset_id": None, "note": "Chưa có dataset active."}
    present = bb.metrics.available_reports(ds)
    dark = sorted({rt for rt in _KPI_SOURCE.values() if rt not in present})
    dark_kpis = sorted([k for k, rt in _KPI_SOURCE.items() if rt not in present])
    return {"dataset_id": ds, "present_report_types": sorted(present),
            "missing_report_types": dark, "dark_kpis": dark_kpis}


def _available() -> list:
    if not os.path.exists(AVAILABLE_META):
        return []
    try:
        with open(AVAILABLE_META, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return []


def pipeline_state() -> dict:
    """1 view hợp nhất: mỗi file -> collected(available) / received(indexed) / ingested."""
    avail = _available()
    cat = {e["file"]: e for e in SC.search()}
    files = {}
    for a in avail:
        fn = a.get("fileName")
        if not fn:
            continue
        c = cat.get(fn)
        files[fn] = {"company": a.get("company"), "collected": True,
                     "received": c is not None, "ingested": bool(c and c.get("ingested"))}
    # file có trong catalog nhưng không trong available (nhận trực tiếp)
    for fn, c in cat.items():
        if fn not in files:
            files[fn] = {"company": c.get("company"), "collected": False,
                         "received": True, "ingested": bool(c.get("ingested"))}
    total = len(files)
    received = sum(1 for f in files.values() if f["received"])
    ingested = sum(1 for f in files.values() if f["ingested"])
    return {"total": total, "collected": len(avail), "received": received, "ingested": ingested,
            "pending_request": total - received, "pending_ingest": received - ingested,
            "files": files}


def status(dataset_id: str = None) -> dict:
    return {"uningested_files": uningested_files(),
            "screens_missing_source": screens_missing_source(dataset_id),
            "pipeline": {k: v for k, v in pipeline_state().items() if k != "files"}}

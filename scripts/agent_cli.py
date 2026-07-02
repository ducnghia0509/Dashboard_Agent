# -*- coding: utf-8 -*-
"""CLI cau noi cho BE (DashBoard_AI) goi qua subprocess -> luon in ra 1 dong JSON tren stdout.
Tach biet tien trinh BE (khong import cheo), chay trong venv Dashboard_Agent (co du deps).

Subcommands:
  profile  <file>                 -> liet ke sheet + canonical_kind_guess + mapping DA HOC (neu co)
  propose  <file> --sheet S       -> goi analyst agent (LLM, qua docker exec) de de xuat SheetMapping
  execute  <file> --sheet S --mapping-file M.json [--period P] [--dry-run]
                                  -> generic_import_execute (ghi GEN_*), tra rows + sample

Vi du:
  .venv/bin/python scripts/agent_cli.py profile /abs/f.xlsx
  .venv/bin/python scripts/agent_cli.py execute /abs/f.xlsx --sheet 131 --mapping-file /tmp/m.json --period 2026-01 --dry-run
"""
import argparse
import json
import os
import subprocess
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(_ROOT, ".env"))


def _out(obj):
    """In dung 1 dong JSON (BE doc dong cuoi)."""
    print(json.dumps(obj, ensure_ascii=False, default=str))


def cmd_profile(args):
    return cmd_plan(args)  # alias — plan đầy đủ hơn (kèm kpi_hints từ guideline)


def _learned_sheet_mappings(sheet: str, canonical_kind: str | None, all_specs: list) -> list:
    """Khớp SheetMapping đã học TỪ danh sách catalog đã nạp sẵn (B3: không đọc đĩa lại mỗi sheet).
    Khớp theo canonical_kind (tái dùng chéo công ty) hoặc đúng tên sheet."""
    seen, out = set(), []
    for r in all_specs:
        sm = r.get("sheet_mapping") or {}
        match = (canonical_kind and sm.get("canonical_kind") == canonical_kind) or sm.get("sheet") == sheet
        fp = r.get("fingerprint")
        if not match or fp in seen:
            continue
        seen.add(fp)
        out.append(sm)
    return out


def _kpi_hints_from_guideline(sheet: str, canonical_kind: str | None, gl_cache: dict) -> list:
    """Khớp sheet/canonical_kind với kpi_glossary.json (sinh từ guideline.xlsx).
    B3: memoize glossary_lookup theo term (nhiều sheet có thể trùng canonical_kind)."""
    from servers import qa_server as qa

    seen_ids, hints = set(), []
    for term in [canonical_kind, sheet]:
        if not term:
            continue
        if term not in gl_cache:
            gl_cache[term] = qa.glossary_lookup(term)
        gl = gl_cache[term]
        for rec in gl.get("kpi_glossary", []):
            rid = rec.get("id")
            if rid in seen_ids:
                continue
            seen_ids.add(rid)
            hints.append({
                "id": rid,
                "chi_tieu": rec.get("chi_tieu"),
                "nguon_du_lieu": rec.get("nguon_du_lieu"),
                "needs_followup": rec.get("needs_followup", False),
            })
    return hints[:8]


def _dashboard_relevance(learned: list, kpi_hints: list, canonical_kind: str | None) -> str:
    """wired = khớp guideline hoặc đã học; optional = có loại chuẩn nhưng chưa có KPI; unused = còn lại."""
    if learned or kpi_hints:
        return "wired"
    if canonical_kind:
        return "optional"
    return "unused"


def cmd_plan(args):
    """Profile + khớp guideline (kpi_glossary) + mapping đã học — không cần LLM.

    B1/B2: chỉ mở workbook 1 lần qua sheet_profile(sheet=None) — nhẹ (chỉ sheetnames +
    canonical_kind_guess). BỎ template_analyze (parse thứ 2): plan chỉ chạy cho template LẠ
    (handoff /import đã lọc file khớp 9 báo cáo / Tháng) nên fixed_report_type luôn null ở đây.
    B3: nạp catalog report_specs + cache glossary 1 lần, khớp trong bộ nhớ cho mọi sheet.
    """
    from servers import ingest_server as ing
    from servers.common import memory

    prof = ing.sheet_profile(args.file)          # 1 lần mở workbook
    all_specs = memory.report_spec_search()      # 1 lần đọc catalog
    gl_cache: dict = {}                          # memoize glossary theo term
    sheets = []
    counts = {"wired": 0, "optional": 0, "unused": 0}
    for s in prof.get("sheets", []):
        name = s.get("sheet")
        ck = s.get("canonical_kind_guess")
        learned = _learned_sheet_mappings(name, ck, all_specs)
        kpi_hints = _kpi_hints_from_guideline(name, ck, gl_cache)
        rel = _dashboard_relevance(learned, kpi_hints, ck)
        counts[rel] = counts.get(rel, 0) + 1
        sheets.append({
            "sheet": name,
            "canonical_kind_guess": ck,
            "learned_mappings": learned,
            "kpi_hints": kpi_hints,
            "guideline_match": bool(kpi_hints),
            "ready": bool(learned),
            "dashboard_relevance": rel,
        })
    _out({
        "ok": True,
        "file_name": prof.get("file_name"),
        "fixed_report_type": None,  # plan chỉ cho template lạ; handoff đã lọc template chuẩn
        "sheets": sheets,
        "sheet_summary": counts,
        "guideline_source": "DashBoard_AI/guideline.xlsx → kpi_glossary.json",
    })


def _analyst_propose_prompt(file_path: str, sheet: str) -> str:
    return f"""Bạn là subagent analyst (SKILL: Dashboard_Agent/agents/analyst/SKILL.md).

Phân tích file Excel và TRẢ VỀ DUY NHẤT 1 khối ```json SheetMapping cho sheet '{sheet}'.

File: {file_path}
Sheet: {sheet}

Quy trình BẮT BUỘC (dùng MCP tools, đọc guideline qua glossary — KHÔNG hard-code tên cột):
1. sheet_profile(file_path='{file_path}', sheet='{sheet}') — đọc row_sample/col_sample, phát hiện header 1-2 dòng, orientation row_major/column_major.
2. report_spec_search(sheet='{sheet}') và report_spec_search(canonical_kind=...) nếu sheet_profile/canonical đoán được loại (TK131, TK331, KQKD, CDKT, TSCD...).
3. glossary_lookup(term=<tên sheet hoặc canonical_kind>) — khớp kpi_glossary.json (nguồn từ guideline.xlsx): xem nguon_du_lieu, chi_tieu, công thức để biết sheet phục vụ chỉ số nào.
4. Đối chiếu display_contract.json — field/màn hình FE nào cần dữ liệu từ sheet này.
5. Dựng SheetMapping: orientation, data_start_row (0-based), columns/entities+row_roles, target_report_type (tiền tố GEN_), canonical_kind, matched_kpi_ids.

Chỉ số 0-based. role: entity_code|entity_name|label|skip|<do_luong>.
KHÔNG giải thích ngoài khối JSON."""


def _propose_mapping(file_path: str, sheet: str, timeout: int = 240):
    """Goi analyst agent (LLM) qua docker exec -> tra (mapping|None, reply_text)."""
    prompt = _analyst_propose_prompt(file_path, sheet)
    cmd = ["docker", "exec", "openclaw", "openclaw", "agent", "--agent", "analyst",
           "--json", "--timeout", str(timeout), "-m", prompt]
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout + 30)
    except subprocess.TimeoutExpired:
        return None, "analyst agent timeout"
    reply = _extract_agent_text(p.stdout)
    mapping = _extract_json_block(reply) if reply else None
    return mapping, reply


def cmd_propose(args):
    mapping, reply = _propose_mapping(args.file, args.sheet, args.timeout)
    _out({"ok": bool(mapping), "proposed_mapping": mapping, "agent_text": reply})


def cmd_autobatch(args):
    """Nap TAT CA sheet wired trong 1 lan: mapping da hoc dung luon; chua hoc thi propose (neu
    --propose) qua analyst LLM. Moi sheet: generic_import_execute. Tra ket qua tung sheet.
    LUU Y: sheet chua hoc phu thuoc chat luong propose (model yeu co the fail -> status=need_manual)."""
    from servers import ingest_server as ing
    from servers.common import memory

    prof = ing.sheet_profile(args.file)
    all_specs = memory.report_spec_search()
    gl_cache: dict = {}
    results = []
    for s in prof.get("sheets", []):
        name = s.get("sheet")
        ck = s.get("canonical_kind_guess")
        learned = _learned_sheet_mappings(name, ck, all_specs)
        kpi_hints = _kpi_hints_from_guideline(name, ck, gl_cache)
        if _dashboard_relevance(learned, kpi_hints, ck) != "wired":
            continue
        mapping = learned[0] if learned else None
        source = "learned" if mapping else None
        if mapping is None and args.propose:
            mapping, _ = _propose_mapping(args.file, name, args.timeout)
            source = "proposed" if mapping else None
        if mapping is None:
            results.append({"sheet": name, "canonical_kind": ck, "status": "need_manual",
                            "reason": "chua co mapping (bat --propose hoac lam thu cong)"})
            continue
        try:
            r = ing.generic_import_execute(args.file, sheet=name, mapping=mapping,
                                           period=args.period, cong_ty=args.cong_ty, dry_run=args.dry_run)
            results.append({"sheet": name, "canonical_kind": ck,
                            "status": "preview" if args.dry_run else ("duplicate" if r.get("skipped_duplicate") else "imported"),
                            "mapping_source": source, "row_count": r.get("row_count"),
                            "target_report_type": r.get("target_report_type"), "dataset_id": r.get("dataset_id")})
        except Exception as e:
            results.append({"sheet": name, "canonical_kind": ck, "status": "error", "error": f"{type(e).__name__}: {e}"})
    imported = sum(1 for r in results if r["status"] in ("imported", "preview"))
    _out({"ok": True, "file_name": prof.get("file_name"), "wired_total": len(results),
          "imported": imported, "results": results})


def cmd_execute(args):
    from servers import ingest_server as ing

    with open(args.mapping_file, encoding="utf-8") as fh:
        mapping = json.load(fh)
    r = ing.generic_import_execute(
        args.file, sheet=args.sheet, mapping=mapping, period=args.period,
        cong_ty=args.cong_ty, dry_run=args.dry_run,
    )
    r["ok"] = True
    _out(r)


def _extract_agent_text(stdout: str):
    try:
        env = json.loads(stdout)
    except json.JSONDecodeError:
        return None
    # tim finalAssistantVisibleText o bat ky do sau nao
    stack = [env]
    while stack:
        node = stack.pop()
        if isinstance(node, dict):
            for k in ("finalAssistantVisibleText", "finalAssistantRawText"):
                if isinstance(node.get(k), str) and node[k].strip():
                    return node[k]
            stack.extend(node.values())
        elif isinstance(node, list):
            stack.extend(node)
    return None


def _extract_json_block(text: str):
    import re
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    raw = m.group(1) if m else None
    if not raw:
        # thu lay khoi { ... } lon nhat
        i, j = text.find("{"), text.rfind("}")
        raw = text[i:j + 1] if 0 <= i < j else None
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("profile"); p.add_argument("file"); p.set_defaults(fn=cmd_profile)
    p = sub.add_parser("plan"); p.add_argument("file"); p.set_defaults(fn=cmd_plan)
    p = sub.add_parser("propose"); p.add_argument("file"); p.add_argument("--sheet", required=True)
    p.add_argument("--timeout", type=int, default=240); p.set_defaults(fn=cmd_propose)
    p = sub.add_parser("execute"); p.add_argument("file"); p.add_argument("--sheet", required=True)
    p.add_argument("--mapping-file", required=True); p.add_argument("--period"); p.add_argument("--cong-ty", dest="cong_ty")
    p.add_argument("--dry-run", action="store_true"); p.set_defaults(fn=cmd_execute)
    p = sub.add_parser("autobatch"); p.add_argument("file"); p.add_argument("--period"); p.add_argument("--cong-ty", dest="cong_ty")
    p.add_argument("--propose", action="store_true"); p.add_argument("--dry-run", action="store_true")
    p.add_argument("--timeout", type=int, default=240); p.set_defaults(fn=cmd_autobatch)

    args = ap.parse_args()
    try:
        args.fn(args)
    except Exception as e:
        import traceback
        _out({"ok": False, "error": f"{type(e).__name__}: {e}", "trace": traceback.format_exc()[-800:]})
        sys.exit(1)


if __name__ == "__main__":
    main()

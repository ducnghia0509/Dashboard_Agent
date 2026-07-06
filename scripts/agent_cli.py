# -*- coding: utf-8 -*-
"""CLI cau noi cho BE (DashBoard_AI) goi qua subprocess -> luon in ra 1 dong JSON tren stdout.
Tach biet tien trinh BE (khong import cheo), chay trong venv Dashboard_Agent (co du deps).

Subcommands:
  plan/profile <file>             -> liet ke sheet + canonical_kind_guess + mapping DA HOC +
                                     de xuat scope/basis/man hinh (chua LLM, deterministic)
  propose  <file> --sheet S       -> heuristic (TK131/TK331) hoac goi analyst agent (LLM)
  execute  <file> --sheet S --mapping-file M.json [--period P] [--dry-run]
                                  -> generic_import_execute (ghi GEN_*), tra rows + sample
  autobatch <file> [--propose] [--dry-run] -> nap tat ca sheet wired 1 lan
  confirm <canonical_kind> --scope --basis --target-screen --chi-tieu
                                  -> luu quyet dinh phan loai (1 lan, dung lai cho moi file)

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
    Khớp theo canonical_kind (tái dùng chéo công ty) hoặc đúng tên sheet. Ưu tiên bản đã
    verified=True (ghi thật thành công) lên đầu — bản chỉ qua dry-run (verified=False) vẫn
    dùng được (khép vòng học sớm) nhưng đáng tin cậy thấp hơn. Trả nguyên SheetMapping (không
    nhét thêm field verified vào — tránh lẫn với schema mapping thật khi dùng lại để execute)."""
    seen, candidates = set(), []
    for r in all_specs:
        sm = r.get("sheet_mapping") or {}
        match = (canonical_kind and sm.get("canonical_kind") == canonical_kind) or sm.get("sheet") == sheet
        fp = r.get("fingerprint")
        if not match or fp in seen:
            continue
        seen.add(fp)
        candidates.append((bool(r.get("verified", True)), sm))
    candidates.sort(key=lambda t: t[0], reverse=True)
    return [sm for _verified, sm in candidates]


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
                "nhom_bao_cao": rec.get("nhom_bao_cao"),
                "nhom_con": rec.get("nhom_con"),
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
    from servers.common import canonical, memory

    routed = ing.sheet_routes(args.file)         # 1 lần mở workbook — route theo TÊN + NỘI DUNG
    all_specs = memory.report_spec_search()      # 1 lần đọc catalog
    gl_cache: dict = {}                          # memoize glossary theo term
    file_name = routed.get("file_name") or os.path.basename(args.file)
    scope_from_filename = canonical.guess_scope(file_name)  # rẻ — chỉ so tên file, không mở sheet
    sheets = []
    counts = {"wired": 0, "optional": 0, "unused": 0}
    unrecognized = []   # KHÔNG drop im lặng: sheet có dữ liệu mà chưa route được -> cần analyst/người
    for r in routed.get("routes", []):
        name = r["sheet"]
        ck = r.get("canonical_kind")             # đã ưu tiên tên, fallback nội dung
        target_sheet = r.get("target_sheet")
        if r.get("status") == "unknown":
            unrecognized.append(name)
        learned = _learned_sheet_mappings(name, ck, all_specs)
        kpi_hints = _kpi_hints_from_guideline(name, ck, gl_cache)
        rel = _dashboard_relevance(learned, kpi_hints, ck)
        counts[rel] = counts.get(rel, 0) + 1
        # Đề xuất chỉ tiêu/màn hình phục vụ (đọc guideline, KHÔNG tự ghi số vào KPI nào —
        # chỉ để hiện gợi ý cho người dùng xác nhận, xem AgentImportPanel + /import/agent/confirm).
        proposed_kpi = kpi_hints[0] if kpi_hints else None
        proposed_screen = (canonical.guess_screen(proposed_kpi.get("nhom_bao_cao", ""),
                                                  proposed_kpi.get("nhom_con", ""))
                          if proposed_kpi else None)
        binding = memory.binding_get(ck) if ck else None  # đã xác nhận từ trước (công ty khác/lần trước)
        sheets.append({
            "sheet": name,
            "canonical_kind_guess": ck,
            "target_sheet": target_sheet,
            "route_via": r.get("route_via"),
            "route_status": r.get("status"),
            "learned_mappings": learned,
            "kpi_hints": kpi_hints,
            "guideline_match": bool(kpi_hints),
            "ready": bool(learned),
            "dashboard_relevance": rel,
            "scope_guess": scope_from_filename,
            "proposed_kpi": proposed_kpi,
            "proposed_screen": proposed_screen,
            "binding": binding,
        })
    _out({
        "ok": True,
        "file_name": file_name,
        "fixed_report_type": None,  # plan chỉ cho template lạ; handoff đã lọc template chuẩn
        "sheets": sheets,
        "sheet_summary": counts,
        "unrecognized_sheets": unrecognized,   # sheet lạ có dữ liệu — hiển thị để không bỏ sót
        "guideline_source": "DashBoard_AI/guideline.xlsx → kpi_glossary.json",
    })


# canonical_kind mà heuristic tất định (không LLM) áp được — sổ tổng hợp công nợ dạng
# header 2 tầng (Đầu kỳ/Phát sinh/Cuối kỳ) x (Nợ/Có) + cột Mã ĐT/Tên ĐT. Đã verify đúng số
# liệu (KH 02010001) trên file thật cho TK131; TK331 cùng mẫu S31-DN nên cấu trúc giống hệt.
_HEURISTIC_CANONICAL_KINDS = {"TK131", "TK331"}


def _forward_fill(cells: list) -> list:
    out, last = [], None
    for c in cells:
        if c not in (None, ""):
            last = c
        out.append(last)
    return out


def _heuristic_tk_mapping(file_path: str, sheet: str, canonical_kind: str):
    """Heuristic TẤT ĐỊNH (không LLM) cho sheet dạng 'sổ tổng hợp công nợ' TK131/TK331:
    dò dòng phụ đề 'Nợ/Có', dòng nhóm kỳ ngay phía trên ('Đầu kỳ'/'Phát sinh'/'Cuối kỳ',
    forward-fill do cell merge), và 2 cột 'Mã ĐT'/'Tên ĐT'. Trả (None, {}) nếu không khớp
    cấu trúc mong đợi (để cmd_propose/cmd_autobatch tự fallback sang LLM) — KHÔNG đoán liều.

    Trả (mapping, meta): meta gồm basis_guess ('luyke'|'thang'|None) suy từ chính text tiêu
    đề đã đọc được trong row_sample (KHÔNG mở thêm file — tận dụng lại lượt đọc này)."""
    from servers import ingest_server as ing
    from servers.common import be_bridge as bb
    from servers.common import canonical

    prof = ing.sheet_profile(file_path, sheet=sheet, max_rows=20, max_cols=15)
    rows = prof.get("row_sample", [])
    norm = lambda v: bb.normalize_header(v, True)  # noqa: E731

    basis_guess = None
    for row in rows[:10]:
        for cell in row:
            if isinstance(cell, str):
                basis_guess = canonical.guess_basis(cell)
                if basis_guess:
                    break
        if basis_guess:
            break

    sub_idx = None
    for i, row in enumerate(rows):
        hits = sum(1 for c in row if norm(c).startswith("no") or norm(c).startswith("co"))
        if hits >= 2:
            sub_idx = i
            break
    if sub_idx is None or sub_idx == 0:
        return None, {"basis_guess": basis_guess}

    group_row = _forward_fill(rows[sub_idx - 1])
    sub_row = rows[sub_idx]
    code_idx = name_idx = None
    for idx, cell in enumerate(rows[sub_idx - 1]):
        t = norm(cell)
        if code_idx is None and t.startswith("ma "):
            code_idx = idx
        elif name_idx is None and t.startswith("ten "):
            name_idx = idx
    if code_idx is None or name_idx is None:
        return None, {"basis_guess": basis_guess}

    columns = [{"index": code_idx, "role": "entity_code"}, {"index": name_idx, "role": "entity_name"}]
    for idx in range(len(sub_row)):
        if idx in (code_idx, name_idx):
            continue
        st = norm(sub_row[idx])
        if not (st.startswith("no") or st.startswith("co")):
            continue
        gt = norm(group_row[idx]) if idx < len(group_row) else ""
        phase = ("dau_ky" if "dau ky" in gt else "phat_sinh" if "phat sinh" in gt
                else "cuoi_ky" if "cuoi ky" in gt else None)
        if not phase:
            continue
        side = "no" if st.startswith("no") else "co"
        columns.append({"index": idx, "role": f"{phase}_{side}"})
    if len(columns) <= 2:  # không tìm được cột đo lường nào -> không tin cấu trúc này
        return None, {"basis_guess": basis_guess}

    mapping = {
        "orientation": "row_major", "data_start_row": sub_idx + 1,
        "target_report_type": f"GEN_{canonical_kind}", "canonical_kind": canonical_kind,
        "columns": columns, "header_rows": [sub_idx - 1, sub_idx],
        "notes": ["heuristic tất định: header 2 tầng (kỳ) x (Nợ/Có) — không qua LLM"],
    }
    return mapping, {"basis_guess": basis_guess}


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
    from servers.common import canonical

    ck = canonical.guess_canonical_kind(args.sheet)
    scope_guess = canonical.guess_scope(os.path.basename(args.file))
    if ck in _HEURISTIC_CANONICAL_KINDS:
        mapping, meta = _heuristic_tk_mapping(args.file, args.sheet, ck)
        if mapping:
            _out({"ok": True, "proposed_mapping": mapping, "agent_text": None, "source": "heuristic",
                  "scope_guess": scope_guess, "basis_guess": meta.get("basis_guess")})
            return
        # heuristic không khớp cấu trúc mong đợi -> fallback LLM như bình thường bên dưới.
    mapping, reply = _propose_mapping(args.file, args.sheet, args.timeout)
    _out({"ok": bool(mapping), "proposed_mapping": mapping, "agent_text": reply, "source": "llm",
          "scope_guess": scope_guess, "basis_guess": None})


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
        if mapping is None and ck in _HEURISTIC_CANONICAL_KINDS:
            mapping, _heur_meta = _heuristic_tk_mapping(args.file, name, ck)
            source = "heuristic" if mapping else None
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


def cmd_confirm(args):
    """Lưu 1 quyết định phân loại (scope/basis/màn hình) do NGƯỜI xác nhận cho 1 canonical_kind.
    Dùng lại cho MỌI file/công ty cùng loại báo cáo sau này (agent_bridge /import/agent/confirm).
    KHÔNG tự ghi số vào KPI nào — chỉ là metadata phân loại, xem servers/common/memory.py."""
    from servers.common import memory

    rec = memory.binding_save(args.canonical_kind, scope=args.scope, basis=args.basis,
                              target_screen=args.target_screen, chi_tieu=args.chi_tieu,
                              kpi_id=args.kpi_id)
    _out({"ok": True, "binding": rec})


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


def cmd_autofill(args):
    """LOOP TẤT ĐỊNH điền template vàng — thay 1-prompt-nhồi-tất-cả. Mỗi sheet rơi vào ĐÚNG 1
    'bucket' (đảm bảo phủ, không sót): filled_learned (đã học mapping -> điền ngay, KHÔNG LLM) |
    need_llm (đã route ra sheet đích nhưng chưa học -> analyst đề xuất) | need_human (chưa nhận
    diện được loại) | skip (metadata) | empty. dry_run=true chỉ preview, không ghi DB."""
    from servers import ingest_server as ing
    from servers import template_filler as tf
    from servers.common import memory

    routes = ing.sheet_routes(args.file).get("routes", [])
    with open(args.file, "rb") as fh:
        data = fh.read()
    fname = os.path.basename(args.file)
    period = args.period or tf.guess_period(fname)
    headers = tf._all_sheet_headers(data)   # mở workbook 1 LẦN (file nặng load chậm)
    ledger = []
    for r in routes:
        sheet, status, target = r["sheet"], r["status"], r.get("target_sheet")
        if status in ("skip_metadata", "empty"):
            ledger.append({"sheet": sheet, "bucket": status.replace("skip_metadata", "skip"),
                           "target_sheet": None})
            continue
        if status == "unknown":
            ledger.append({"sheet": sheet, "bucket": "need_human", "target_sheet": None,
                           "reason": "chưa nhận diện được loại báo cáo"})
            continue
        head = headers.get(sheet) or []
        hr = tf._header_row_of(head)
        cols = head[hr] if hr < len(head) else []
        spec = memory.fill_spec_find(memory.source_fingerprint(sheet, cols))
        if not spec:
            ledger.append({"sheet": sheet, "bucket": "need_llm", "target_sheet": target,
                           "canonical_kind": r.get("canonical_kind"), "route_via": r.get("route_via"),
                           "reason": "chưa học mapping — cần analyst đề xuất"})
            continue
        res = tf.fill_from_source(data, sheet, spec["target_sheet"], spec["mapping"],
                                  period=period, cong_ty=args.cong_ty, file_name=fname,
                                  dry_run=args.dry_run, auto_import=not args.dry_run, learn=False,
                                  value_scale=spec.get("value_scale", 1.0), constants=spec.get("constants"),
                                  rename_rows=spec.get("rename_rows"), source_path=args.file)
        ledger.append({"sheet": sheet, "bucket": "filled_learned", "target_sheet": spec["target_sheet"],
                       "dry_run": args.dry_run, "row_count": res.get("row_count"),
                       "rows_written": res.get("rows_written"),
                       "imported": (res.get("import") or {}).get("ok")})
    summary = {}
    for e in ledger:
        summary[e["bucket"]] = summary.get(e["bucket"], 0) + 1
    _out({"ok": True, "file": fname, "period": period, "summary": summary, "ledger": ledger})


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
    p = sub.add_parser("autofill"); p.add_argument("file"); p.add_argument("--period")
    p.add_argument("--cong-ty", dest="cong_ty"); p.add_argument("--dry-run", action="store_true")
    p.set_defaults(fn=cmd_autofill)
    p = sub.add_parser("confirm"); p.add_argument("canonical_kind")
    p.add_argument("--scope"); p.add_argument("--basis"); p.add_argument("--target-screen", dest="target_screen")
    p.add_argument("--chi-tieu", dest="chi_tieu"); p.add_argument("--kpi-id", dest="kpi_id")
    p.set_defaults(fn=cmd_confirm)

    args = ap.parse_args()
    try:
        args.fn(args)
    except Exception as e:
        import traceback
        _out({"ok": False, "error": f"{type(e).__name__}: {e}", "trace": traceback.format_exc()[-800:]})
        sys.exit(1)


if __name__ == "__main__":
    main()

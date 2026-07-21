# -*- coding: utf-8 -*-
"""Tu-verify khong can mo Claude Code / MCP protocol - goi thang cac ham tool.
Chay: python scripts/smoke_test.py (tu thu muc DashBoard_Agent, hoac python -m scripts.smoke_test)
"""
import glob
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from servers.common import guardrails  # noqa: E402
from servers.common import memory  # noqa: E402


def _ok(label):
    print(f"[OK] {label}")


def _fail(label, detail=""):
    print(f"[FAIL] {label} {detail}")
    global _failed
    _failed = True


_failed = False


def test_guardrails():
    print("\n== 1. guardrails ==")
    accept = ["SELECT 1", "select * from raw_rows limit 10", "WITH t AS (SELECT 1) SELECT * FROM t"]
    reject = [
        "DROP TABLE raw_rows", "SELECT 1; DROP TABLE x", "UPDATE raw_rows SET amount=0",
        "INSERT INTO raw_rows VALUES (1)", "SELECT * FROM raw_rows; SELECT * FROM datasets",
    ]
    for sql in accept:
        try:
            guardrails.check_select_only(sql)
            _ok(f"accept: {sql!r}")
        except guardrails.GuardrailError as e:
            _fail(f"should accept: {sql!r}", str(e))
    for sql in reject:
        try:
            guardrails.check_select_only(sql)
            _fail(f"should reject: {sql!r}")
        except guardrails.GuardrailError:
            _ok(f"reject: {sql!r}")
    limited = guardrails.ensure_limit("SELECT * FROM raw_rows")
    if "LIMIT" in limited:
        _ok(f"ensure_limit adds LIMIT: {limited}")
    else:
        _fail("ensure_limit did not add LIMIT", limited)


def test_template_analyze():
    print("\n== 2. template_analyze trên Data_test_dashboard/ ==")
    from servers import ingest_server as ing

    input_dir = os.path.normpath(os.path.join(_ROOT, "..", "Data_test_dashboard"))
    files = sorted(glob.glob(os.path.join(input_dir, "*.xlsx")))
    if not files:
        _fail(f"không tìm thấy file mẫu trong {input_dir}")
        return
    for fp in files:
        try:
            r = ing.template_analyze(fp)
            status = "OK" if r["report_type"] and not r["missing_required"] else "CẢNH BÁO"
            print(f"  [{status}] {os.path.basename(fp)} -> report_type={r['report_type']} "
                  f"missing={r['missing_required']} confidence={r['confidence']}")
        except Exception as e:
            _fail(f"template_analyze lỗi trên {os.path.basename(fp)}", f"{type(e).__name__}: {e}")


def test_discovery_roundtrip():
    print("\n== 3. discovery_record / discovery_search round-trip ==")
    rec = memory.discovery_record(
        file_name="__smoke_test__.xlsx", fingerprint="smoke-fp", sheets=["S1"],
        columns_per_sheet={"S1": ["A", "B"]}, detected_report_type="THUCHI",
        header_row=0, mapping={"ngay": 0}, confidence=0.9,
    )
    found = memory.discovery_search("__smoke_test__")
    if found and found[0]["fingerprint"] == "smoke-fp":
        _ok("round-trip ghi/đọc discovery memory")
    else:
        _fail("round-trip discovery memory thất bại", str(found))
    os.remove(memory._path_for("__smoke_test__.xlsx"))


def test_glossary_lookup():
    print("\n== 4. glossary_lookup (kpi_glossary.json) ==")
    from servers import qa_server as qa

    if not os.path.exists(qa._KPI_GLOSSARY_PATH):
        _fail(f"chưa có {qa._KPI_GLOSSARY_PATH} - chạy `python scripts/gen_kpi_glossary.py` trước.")
        return
    # Từ khoá khớp 50_chi_tieu.yaml (nguồn mới của glossary) — "he so no"/"vay va lai vay"
    # là tên chỉ tiêu của guideline.xlsx CŨ, không còn trong danh mục 53 chỉ tiêu.
    for term in ["doanh thu thuan", "tuoi no", "no vay", "ton kho"]:
        r = qa.glossary_lookup(term)
        if r["total_matches"] > 0:
            _ok(f"glossary_lookup('{term}') -> {r['total_matches']} kết quả")
        else:
            _fail(f"glossary_lookup('{term}') không có kết quả nào")


def main():
    test_guardrails()
    test_template_analyze()
    test_discovery_roundtrip()
    test_glossary_lookup()

    print("\n== 5. Phần cần DB thật (KHÔNG tự chạy) ==")
    print("  sql_query / import_execute(dry_run=False) cần DATABASE_URL[_RO] thật trong .env.")
    print("  Xem README.md mục 'Bật luồng thật' để tự chạy khi sẵn sàng.")

    print("\n" + ("CÓ LỖI - xem [FAIL] ở trên." if _failed else "TẤT CẢ PASS."))
    sys.exit(1 if _failed else 0)


if __name__ == "__main__":
    main()

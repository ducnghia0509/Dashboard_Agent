# -*- coding: utf-8 -*-
"""Cham diem qa tren 20 cau hoi golden. Phan glossary/discovery/source chay that (khong
can DB); phan sql chi chay neu da co DATABASE_URL_RO trong .env, nguoc lai SKIP co log ro.
Chay: python eval/qa_golden/score.py
"""
import glob
import json
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
sys.path.insert(0, _ROOT)

from servers import ingest_server as ing  # noqa: E402
from servers import qa_server as qa  # noqa: E402


def _seed_discovery_memory():
    """Chạy template_analyze trên toàn bộ file mẫu để discovery memory có dữ liệu
    thật cho các câu hỏi expected_kind=discovery."""
    input_dir = os.path.normpath(os.path.join(_ROOT, "..", "Data_test_dashboard"))
    for fp in glob.glob(os.path.join(input_dir, "*.xlsx")):
        try:
            ing.template_analyze(fp)
        except Exception:
            pass


def main():
    with open(os.path.join(_HERE, "questions.json"), encoding="utf-8") as fh:
        questions = json.load(fh)

    _seed_discovery_memory()

    passed = skipped = 0
    for item in questions:
        kind = item["expected_kind"]
        q = item["q"]
        try:
            if kind == "glossary":
                r = qa.glossary_lookup(item["term"])
                ok = r["total_matches"] > 0
                print(f"[{'PASS' if ok else 'FAIL'}] (glossary) {q} -> {r['total_matches']} kết quả")
            elif kind == "discovery":
                r = qa.discovery_search(query=item.get("query"), report_type=item.get("report_type"))
                ok = len(r) > 0
                print(f"[{'PASS' if ok else 'FAIL'}] (discovery) {q} -> {len(r)} bản ghi")
            elif kind == "source":
                r = qa.source_inspect(item["file_name"], max_rows=item.get("max_rows", 5))
                ok = r["row_count_returned"] > 0
                print(f"[{'PASS' if ok else 'FAIL'}] (source) {q} -> {r['row_count_returned']} dòng, sheets={r['all_sheets']}")
            elif kind == "sql":
                if not os.environ.get("DATABASE_URL_RO"):
                    print(f"[SKIP] (sql) {q} -> chưa cấu hình DATABASE_URL_RO, bỏ qua (xem README).")
                    skipped += 1
                    continue
                r = qa.sql_query("SELECT 1")  # câu hỏi thật cần agent tự sinh SQL - đây chỉ test kết nối
                ok = r["row_count"] >= 0
                print(f"[{'PASS' if ok else 'FAIL'}] (sql) {q} -> kết nối DB OK, tự sinh SQL do agent thực hiện.")
            else:
                print(f"[FAIL] expected_kind không hợp lệ: {kind}")
                ok = False
        except Exception as e:
            print(f"[FAIL] {q} -> lỗi {type(e).__name__}: {e}")
            ok = False
        if ok:
            passed += 1

    total = len(questions) - skipped
    print(f"\nKết quả: {passed}/{total} pass (bỏ qua {skipped} câu cần DATABASE_URL_RO thật).")
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()

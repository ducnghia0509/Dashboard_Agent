# -*- coding: utf-8 -*-
"""Cham diem analyst tren file mau: chay template_analyze tren tung file trong
Data_test_dashboard/, so report_type + missing_required voi expected.json.
Chay: python eval/template/score.py
"""
import json
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
sys.path.insert(0, _ROOT)

from servers import ingest_server as ing  # noqa: E402


def main():
    with open(os.path.join(_HERE, "expected.json"), encoding="utf-8") as fh:
        expected = json.load(fh)

    input_dir = os.path.normpath(os.path.join(_ROOT, "..", "Data_test_dashboard"))
    passed, total = 0, len(expected)
    for exp in expected:
        fp = os.path.join(input_dir, exp["file_name"])
        if not os.path.exists(fp):
            print(f"[SKIP] không tìm thấy file: {exp['file_name']}")
            total -= 1
            continue
        r = ing.template_analyze(fp)
        rt_ok = r["report_type"] == exp["expected_report_type"]
        covered_ok = (not r["missing_required"]) == exp["required_covered"]
        ok = rt_ok and covered_ok
        status = "PASS" if ok else "FAIL"
        print(f"[{status}] {exp['file_name']}: report_type={r['report_type']} "
              f"(kỳ vọng {exp['expected_report_type']}), missing_required={r['missing_required']}")
        if ok:
            passed += 1

    print(f"\nKết quả: {passed}/{total} pass ({100 * passed / total if total else 0:.0f}%)")
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Dựng ĐÚNG số mà dashboard hiển thị cho một màn (gọi chính builder của API, không qua HTTP/đăng nhập).

Chạy bằng venv của tc-admin-api, DB chỉ đọc:  DATABASE_URL=... PGOPTIONS="-c default_transaction_read_only=on" \
  PYTHONPATH=<tc-admin-api> python man_hien_thi_dump.py '{"man":"vhkd0","period":"2026-09"}'  -> JSON ra stdout."""
import json
import sys


def main():
    a = json.loads(sys.argv[1])
    from app import repo
    from app.services import metric_service as ms
    did = None
    if a.get("period"):
        m = repo.month_datasets_by_period()
        if a["period"] not in m:
            print(json.dumps({"loi": f"không có kỳ {a['period']} trong DB", "ky_co": sorted(m)}, ensure_ascii=False))
            return
        did = m[a["period"]]["id"]
    res = ms.screen(a["man"], frm=a.get("tu_ngay"), to=a.get("den_ngay"), cong_ty=a.get("cong_ty"), khoi=a.get("khoi"),
                    kind="month" if a.get("grain", "month") == "month" else "day", dataset_id=did, actor=None,
                    grain=a.get("grain", "month"))
    print(json.dumps(res, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()

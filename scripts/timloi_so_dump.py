# -*- coding: utf-8 -*-
"""Dựng số dashboard cho NHIỀU (màn, kỳ) trong MỘT tiến trình — bản nhiều-yêu-cầu của `man_hien_thi_dump.py`, dùng cho
bot tìm lỗi (`servers/common/tim_loi.py`: `tl_so_dashboard`, `tl_so_qua_ky`, `tl_truy_vet`).

Vì sao tách file: import `app.*` của tc-admin-api mất ~1,5 giây; đo một key qua 6 kỳ × 2 DB bằng script cũ là 12 lần
import. Ở đây mỗi DB một tiến trình. Gọi đúng `metric_service.screen` như script cũ (cùng tham số, cùng cách chọn
dataset tháng) — đổi một bên thì đổi cả bên kia.

Chạy bằng venv của tc-admin-api, DB chỉ đọc (PGOPTIONS default_transaction_read_only=on do tiến trình cha đặt):
    PYTHONPATH=<tc-admin-api> python timloi_so_dump.py '{"yeu_cau":[{"man":"thue","period":"2026-01","khoi":"Khối KD Dự án"}]}'
-> JSON list cùng thứ tự, mỗi phần tử là kết quả builder hoặc {"loi": ...}."""
import json
import sys


def main():
    a = json.loads(sys.argv[1])
    from app import repo
    from app.services import metric_service as ms
    m = repo.month_datasets_by_period()
    ra = []
    for y in a.get("yeu_cau") or []:
        try:
            did = None
            if y.get("period"):
                if y["period"] not in m:
                    ra.append({"loi": f"không có kỳ {y['period']} trong DB", "ky_co": sorted(m)})
                    continue
                did = m[y["period"]]["id"]
            g = y.get("grain", "month")
            ra.append(ms.screen(y["man"], frm=y.get("tu_ngay"), to=y.get("den_ngay"), cong_ty=y.get("cong_ty"),
                                khoi=y.get("khoi"), kind="month" if g == "month" else "day", dataset_id=did,
                                actor=None, grain=g))
        except Exception as e:  # noqa: BLE001 — một yêu cầu hỏng không được làm mất kết quả các yêu cầu khác
            ra.append({"loi": f"{type(e).__name__}: {str(e)[:300]}"})
    print(json.dumps(ra, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()

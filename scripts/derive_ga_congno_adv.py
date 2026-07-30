# -*- coding: utf-8 -*-
"""Bổ sung chiều ĐẢO công nợ (PTHU_ADV/PTRA_ADV) riêng cho GA — vì derive_congno_advance.py (bản
chung) không nhận layout sheet 'PThu'/'PTra' của GA (tên sheet viết tắt, thiếu nhãn 'phai thu'/
'phai tra' + dòng 'Tài khoản: 3xx' mà bản chung cần). Tái dùng bộ dò cột 2 tầng của
derive_ga_congno.py (đã verify đúng layout GA), chỉ đổi CHIỀU lấy số: PThu -> dư CÓ (KH trả tiền
trước) -> PTHU_ADV; PTra -> dư NỢ (mình ứng trước NCC) -> PTRA_ADV. Ghi idempotent theo
source_file+period+report_type (xoá bản ADV cũ cùng khoá trước khi insert), giống hệt cơ chế của
derive_congno_advance.py — KHÔNG đụng PTHU/PTRA gốc.

Chạy (dry-run mặc định):
  .venv/bin/python scripts/derive_ga_congno_adv.py <file.xlsx> --period 2026-06
Ghi thật:
  .venv/bin/python scripts/derive_ga_congno_adv.py <file.xlsx> --period 2026-06 --write
"""
import argparse
import json
import os
import sys

import psycopg

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers.common import be_bridge as bb  # noqa: E402
import agent_cli as A  # noqa: E402
from derive_ga_congno import _find_cols, _rows_of  # noqa: E402

DB_URL = (os.environ.get("DATABASE_URL") or os.environ.get("TC_DATABASE_URL")
          or "postgresql://tc:tc@localhost:5433/tc_dashboard")


def _twin_attrs(cur, source_file, period, twin_rt):
    """Lấy (dataset_id, ngay, khoi, cong_ty) để gắn cho dòng ADV mới — ưu tiên dòng report_type
    THUẬN cùng file+kỳ (twin_rt), fallback BẤT KỲ report_type nào khác cùng file+kỳ nếu twin_rt
    KHÔNG có dòng nào (VD tháng đó mọi KH chỉ có dư Có/Nợ đối bên -> derive_ga_congno.extract() xoá
    sạch PTHU/PTRA thuận -> vẫn cần ghi được ADV vì đây LÀ chiều dữ liệu chính của tháng đó)."""
    cur.execute(
        "SELECT dataset_id, ngay, khoi, cong_ty FROM raw_rows "
        "WHERE source_file=%s AND period_month=%s AND report_type=%s LIMIT 1",
        (source_file, period, twin_rt))
    row = cur.fetchone()
    if row:
        return row
    cur.execute(
        "SELECT dataset_id, ngay, khoi, cong_ty FROM raw_rows "
        "WHERE source_file=%s AND period_month=%s LIMIT 1",
        (source_file, period))
    return cur.fetchone()


def derive(path, period, write=False, cong_ty="GA"):
    wb = bb.fast_load_workbook(path, data_only=True, read_only=True)
    names = {s.strip().lower(): s for s in wb.sheetnames}
    src, khoi = A._source_id(path), A._khoi_of(path)
    out = {"file": os.path.basename(path), "period": period, "write": write,
           "source_file": src, "blocks": []}

    conn = psycopg.connect(DB_URL)
    cur = conn.cursor()

    # (sheet, twin report_type thuận, report_type ADV, tk cần lọc, chiều lấy, nhãn)
    JOBS = [("pthu", "PTHU", "PTHU_ADV", "131", "co", "Người mua trả tiền trước"),
            ("ptra", "PTRA", "PTRA_ADV", "331", "no", "Trả trước NCC")]

    for sheet_key, twin_rt, adv_rt, tk_prefix, side, label in JOBS:
        sn = names.get(sheet_key)
        if not sn:
            out["blocks"].append({"sheet": sheet_key, "skip": "không thấy sheet"})
            continue
        rows = [list(r) for r in wb[sn].iter_rows(values_only=True)]
        hs, cols = _find_cols(rows)
        if hs is None:
            out["blocks"].append({"sheet": sheet_key, "skip": "không dò được header"})
            continue
        recs = []
        for ten, tk, dn, dc, pn, pc, cn, cc in _rows_of(rows, hs, cols):
            if not tk.startswith(tk_prefix):
                continue
            val = cc if side == "co" else dn if False else (cn if side == "no" else cc)
            # chiều ĐẢO: PThu (TK131) -> dư CÓ; PTra (TK331) -> dư NỢ
            val = cc if side == "co" else cn
            if not val or abs(val) < 1:
                continue
            recs.append({"ten": ten, "ty": round(val * 1e-9, 9)})
        tong = round(sum(r["ty"] for r in recs), 6)
        blk = {"sheet": sn, "adv_rt": adv_rt, "label": label,
               "so_doi_tuong": len(recs), "tong_ty": tong,
               "top": sorted(recs, key=lambda x: -x["ty"])[:5]}
        if write:
            attrs = _twin_attrs(cur, src, period, twin_rt)
            if not attrs:
                blk["error"] = f"không thấy twin {twin_rt} cùng source_file+period -> BỎ ghi"
                out["blocks"].append(blk)
                continue
            dataset_id, ngay, khoi_db, cong_ty_db = attrs
            cur.execute("DELETE FROM raw_rows WHERE source_file=%s AND period_month=%s AND report_type=%s",
                        (src, period, adv_rt))
            for k, r in enumerate(recs):
                cur.execute(
                    "INSERT INTO raw_rows (dataset_id, report_type, row_index, ngay, cong_ty, khoi, "
                    "cost_center, period_month, amount, amount2, dim1, dim2, dim3, payload, source_file) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                    (dataset_id, adv_rt, 7000000 + k, ngay, cong_ty_db, khoi_db, None, period,
                     r["ty"], None, r["ten"], None, None,
                     json.dumps({"unit": "ty", "nguon": f"TK{tk_prefix} du {side}"}, ensure_ascii=False),
                     src))
            blk["written"] = len(recs)
        out["blocks"].append(blk)

    if write:
        conn.commit()
    conn.close()
    wb.close()
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--cong-ty", dest="cong_ty", default="GA")
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()
    print(json.dumps(derive(a.file, a.period, a.write, a.cong_ty), ensure_ascii=False, indent=2))

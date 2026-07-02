# -*- coding: utf-8 -*-
"""Mo ta schema raw_rows + report_type hop le, de nhet vao prompt cho qa/analyst
(text-to-SQL). Cau truc bang co dinh (xem DashBoard_AI/backend/app/db.py SCHEMA)
nen khong can cache DB - hard-code mo ta o day."""
from . import be_bridge as bb

RAW_ROWS_COLUMNS = [
    ("id", "bigint", "khoá chính, tự tăng"),
    ("dataset_id", "text", "id của lần import (datasets.id) - mỗi lần import = 1 dataset mới"),
    ("report_type", "text", f"1 trong: {', '.join(bb.REPORT_CODES)}"),
    ("row_index", "int", "số thứ tự dòng trong file gốc (để trace ngược)"),
    ("ngay", "text (ISO date)", "ngày phát sinh / ngày báo cáo"),
    ("cong_ty", "text", "tên công ty (dimension)"),
    ("khoi", "text", "tên khối kinh doanh (9 khối, xem master_data.json)"),
    ("cost_center", "text", "bộ phận / cơ sở / depot"),
    ("period_month", "text (YYYY-MM)", "kỳ tháng, dùng cho báo cáo Tháng"),
    ("amount", "double", "số tiền chính (đơn vị đồng)"),
    ("amount2", "double", "số tiền phụ (vd kế hoạch, giá trị nhập mới...)"),
    ("dim1", "text", "chiều phân loại 1 (khác nhau theo report_type - xem FIELD_DEFS)"),
    ("dim2", "text", "chiều phân loại 2"),
    ("dim3", "text", "chiều phân loại 3"),
    ("payload", "text (JSON)", "dữ liệu gốc dạng JSON, dùng để trace nguồn"),
]


def schema_describe() -> str:
    lines = ["Bảng raw_rows(", ]
    for name, typ, desc in RAW_ROWS_COLUMNS:
        lines.append(f"  {name} {typ}  -- {desc}")
    lines.append(")")
    lines.append("")
    lines.append("report_type hợp lệ:")
    for code in bb.REPORT_CODES:
        label = bb.REPORT_LABELS.get(code, code)
        snap = " (snapshot: lấy ngày mới nhất, KHÔNG cộng dồn)" if code in bb.SNAPSHOT else " (flow: cộng dồn theo khoảng ngày)"
        lines.append(f"  {code} = {label}{snap}")
    lines.append("")
    lines.append("Mỗi report_type có field riêng ở FIELD_DEFS (xem glossary_lookup để tra ý nghĩa cột).")
    return "\n".join(lines)

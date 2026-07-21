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
    ("amount", "double", "số tiền chính — ĐƠN VỊ TỶ ĐỒNG (theo template chuẩn). KHÔNG chia 1e9: 3.13 nghĩa là 3.13 tỷ."),
    ("amount2", "double", "số tiền phụ (kế hoạch/giá trị mới...) — cũng ĐƠN VỊ TỶ ĐỒNG"),
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
    lines.append("")
    lines.append("QUAN TRỌNG: amount/amount2 đã ở đơn vị TỶ ĐỒNG. Trả lời trực tiếp SUM(amount), "
                 "TUYỆT ĐỐI KHÔNG chia cho 1e9. Lọc kỳ bằng period_month='YYYY-MM' hoặc ngay.")
    lines.append("")
    # Đồng bộ số QA với số DASHBOARD (metrics._where của DashBoard_AI) — thiếu 2 điều kiện
    # dưới đây là nguyên nhân số QA lệch số trên màn hình:
    lines.append(
        "ĐỂ KHỚP SỐ VỚI DASHBOARD, mọi câu SELECT PHẢI:\n"
        "  1. Loại file đã ẩn:  AND (source_file IS NULL OR source_file NOT IN "
        "(SELECT source_file FROM hidden_files))\n"
        "  2. Scope dataset: raw_rows chứa NHIỀU lần import/kỳ — lọc period_month='YYYY-MM' "
        "VÀ dataset_id (SELECT id FROM datasets WHERE kind=... AND period=...); SUM toàn bảng "
        "không lọc dataset sẽ cộng chồng nhiều lần import.")
    lines.append(
        "KQKD (HQKD/PNLT) — tên chỉ tiêu dim1 CHƯA chuẩn hoá giữa các đơn vị, dashboard dùng "
        "MÃ/PATTERN sau (dùng y hệt để khớp số):\n"
        "  Doanh thu = dim1='1000'; Tổng chi phí = dim1='1047'; LN sau thuế (HQKD) = dim1='1112'; "
        "Thuế TNDN = dim1='1111';\n"
        "  LNST theo dòng khoản mục (PNLT): dim1 ILIKE '%lợi nhuận%sau thu%'; "
        "LNTT: dim1 ILIKE '%lợi nhuận%trước thu%'; Giá vốn: dim1 LIKE 'GIÁ VỐN%'.")
    return "\n".join(lines)

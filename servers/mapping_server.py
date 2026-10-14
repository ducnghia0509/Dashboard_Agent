# -*- coding: utf-8 -*-
"""MCP server `dashboard_mapping` — bot SO KHỚP MAPPING ↔ LOGIC LÊN SỐ (14/10/2026). Xem servers/common/mapping_check.py.

Chạy:  python scripts/serve_http.py mapping 8814        Chỉ ĐỌC (Excel mapping + mã nguồn + DB read_only).
"""
from mcp.server.fastmcp import FastMCP

from .common import mapping_check as mc

mcp = FastMCP("dashboard_mapping")


@mcp.tool()
def map_danh_muc(env: str = "prod", period: str = None) -> dict:
    """Độ phủ mapping của TỪNG màn dashboard: màn nào có/không có tài liệu mapping, builder đọc những report_type
    nào, report_type nào chưa có dữ liệu. Dùng cho câu hỏi 'mapping đã đủ chưa'. `period` = YYYY-MM (tuỳ chọn)."""
    return mc.danh_muc(env, period)


@mcp.tool()
def map_tim_chi_tieu(tu_khoa: str, man: str = None, tai_lieu: str = None) -> dict:
    """Tìm dòng mapping theo tên chỉ tiêu/từ khoá (không phân biệt dấu). Giới hạn bằng `man` (vd vhkd0) hoặc
    `tai_lieu` (QTTC, VHKD, XDV, DUAN, HTXT, ANTAXI, XANHVP, QLTS). Trả công thức + nguồn dữ liệu theo tài liệu."""
    return mc.tim_chi_tieu(tu_khoa, man, tai_lieu)


@mcp.tool()
def map_logic_man(man: str, so_dong_ma: int = 70) -> dict:
    """Logic thật của một màn: hàm builder (file:dòng), các report_type nó đọc, nguồn ngày, và đoạn mã đầu hàm."""
    return mc.logic_man(man, so_dong_ma)


@mcp.tool()
def map_doi_chieu(man: str, tu_khoa: str, env: str = "prod", period: str = None) -> dict:
    """ĐỐI CHIẾU một chỉ tiêu: dòng mapping ↔ builder ↔ spec nạp ↔ dữ liệu trong DB. Trả cờ nghi vấn tất định
    (mapping không nêu file, không spec nạp file đó, builder không đọc nguồn đó, nguồn chưa có dữ liệu) kèm dòng mã
    nguồn chứa từ khoá để đọc công thức."""
    return mc.doi_chieu(man, tu_khoa, env, period)


if __name__ == "__main__":
    mcp.run()

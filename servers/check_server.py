# -*- coding: utf-8 -*-
"""MCP server `dashboard_check` — bot kiểm số liệu (14/10/2026).

Chạy:  python -m servers.check_server           (stdio)
       python scripts/serve_http.py check 8812   (HTTP/SSE cho OpenClaw gateway .253)

Mọi tool đều TẤT ĐỊNH và CHỈ ĐỌC (xem servers/common/so_lieu_check.py). KHÔNG có tool nhận SQL tự do:
bot không tự tính số, chỉ gọi tool rồi diễn giải kết quả có cấu trúc.
"""
from mcp.server.fastmcp import FastMCP

from .common import so_lieu_check as sc

mcp = FastMCP("dashboard_check")


@mcp.tool()
def check_don_vi() -> dict:
    """Danh sách `don_vi` hợp lệ (tiền tố nguồn trong DB) cho ba tool kiểm bên dưới, và hai DB `env`."""
    return {"prod": sc.don_vi_co_trong_db("prod"), "test": sc.don_vi_co_trong_db("test"),
            "env": ["prod", "test"],
            "luu_y": "TEST_SR/TEST_XDV là đường ống nguồn tự động của SR/XDV, không phải môi trường test."}


@mcp.tool()
def check_thang_vs_ngay(don_vi: str, period: str, env: str = "prod") -> dict:
    """So BC THÁNG với Σ BC NGÀY của một đơn vị trong một kỳ (period = YYYY-MM), từng chỉ tiêu
    P&L/doanh thu/chi phí/HQKD. Trả chênh (ngày − tháng), chỉ tiêu lệch, ngày thiếu, gợi ý nguyên nhân."""
    return sc.thang_vs_ngay(don_vi, period, env)


@mcp.tool()
def check_chi_tiet_vs_tong(don_vi: str, ngay: str, rt_chi_tiet: str = "XDV_CN_RO_D", env: str = "prod") -> dict:
    """Số công nợ hiển thị (PTHU_D) có bằng Σ dòng chi tiết nguồn cùng ngày không (ngay = YYYY-MM-DD).
    Chỉ ra dòng bị bỏ (thường là dòng không có tên đối tượng)."""
    return sc.chi_tiet_vs_tong(don_vi, ngay, rt_chi_tiet, env)


@mcp.tool()
def check_nguon_thieu_cham(don_vi: str, period: str, env: str = "prod", thu_muc: str = "baocaohqkdngay") -> dict:
    """File ngày của đơn vị: ngày nào ĐÃ có file, ngày nào CHƯA có file ở nguồn, ngày nào có file mà DB
    chưa có số, và dòng cron lượt cuối. Dùng trả lời 'sao chưa lên số'."""
    return sc.nguon_thieu_cham(don_vi, period, env, thu_muc)


if __name__ == "__main__":
    mcp.run()

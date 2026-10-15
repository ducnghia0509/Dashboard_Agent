# -*- coding: utf-8 -*-
"""MCP server `dashboard_pull` — bot KÉO FILE + NẠP (14/10/2026). Xem servers/common/keo_nap.py.

Chạy:  python scripts/serve_http.py pull 8813
Có tool GHI vào DB; PROD chỉ nạp được khi NGƯỜI YÊU CẦU gõ `XAC NHAN <mã>` (API ghi xác nhận, LLM không tự tạo được);
kỳ năm cũ cần thêm `DUYET <mã>` kèm ảnh email KSNB (API ghi `anh_duyet/<mã>.json`).
"""
from mcp.server.fastmcp import FastMCP

from .common import keo_nap as kn

mcp = FastMCP("dashboard_pull")


@mcp.tool()
def pull_liet_ke_ky(don_vi: str, thu_muc: str) -> dict:
    """BƯỚC 0 (chỉ đọc danh sách nguồn, không kéo/nạp gì). Liệt kê các kỳ YYYY-MM của một nguồn
    `don_vi`/`thu_muc` và tên file từng kỳ, kèm file không suy được kỳ (sẽ bị bỏ). Dùng khi người dùng nói
    "tất cả các tháng / cả thư mục": lấy mọi tên file ở đây rồi gọi pull_lap_ke_hoach(ten_file=[...]),
    tối đa 40 file một kế hoạch (`so_ke_hoach_can` cho biết cần mấy kế hoạch)."""
    return kn.liet_ke_ky(don_vi, thu_muc)


@mcp.tool()
def pull_lap_ke_hoach(don_vi: str, thu_muc: str, ten_file: list[str] = None, ngay_tu: str = None,
                      ngay_den: str = None, period: str = None, da_co_anh_duyet: bool = False) -> dict:
    """BƯỚC 1. Lập kế hoạch kéo+nạp. `don_vi` = ANTAXI/ANKHACHSAN/XDV/TEST_XDV/…, `thu_muc` = loại nguồn
    (vd baocaohqkdngay). Chọn file bằng `ten_file` (danh sách tên) HOẶC `ngay_tu`/`ngay_den` (YYYY-MM-DD,
    ngày trong tên file) HOẶC `period` (YYYY-MM). Chỉ nhận file có trong danh sách nguồn, tối đa 40 file.
    Kỳ thuộc NĂM CŨ (quy định KSNB): vẫn lập được, kết quả có `can_anh_duyet` + `huong_dan_cho_nguoi_dung` —
    chuyển nguyên văn cho người dùng: gửi tin `DUYET <ma_xac_nhan>` KÈM ẢNH email duyệt của KSNB; hệ thống tự ghi
    nhận ảnh, chưa có thì pull_nap từ chối. `da_co_anh_duyet` đã BỎ: truyền vào cũng bị bỏ qua (ghi nhật ký).
    Trả plan_id và ma_xac_nhan (mã dành cho NGƯỜI DÙNG gõ khi muốn lên PROD; chỉ người yêu cầu dùng được)."""
    return kn.lap_ke_hoach(don_vi, thu_muc, ten_file, ngay_tu, ngay_den, period, da_co_anh_duyet)


@mcp.tool()
def pull_keo_ve(plan_id: str) -> dict:
    """BƯỚC 2. Kéo các file của kế hoạch từ máy kế toán về 252 (chờ tối đa 300 giây)."""
    return kn.keo_ve(plan_id)


@mcp.tool()
def pull_nap(plan_id: str, env: str = "test") -> dict:
    """BƯỚC 3/4. Nạp file vào DB và trả bảng so số TRƯỚC/SAU. env='test' làm ngay. env='prod' chỉ chạy khi
    kế hoạch đã nạp TEST thành công VÀ chính người yêu cầu đã gõ 'XAC NHAN <ma_xac_nhan>' (nếu chưa, tool trả
    can_xac_nhan=true kèm hướng dẫn — KHÔNG được tự bịa xác nhận). Kế hoạch có kỳ năm cũ: cả TEST lẫn PROD
    cần ảnh duyệt (`DUYET <ma>` kèm ảnh); thiếu thì trả can_anh_duyet + hướng dẫn, không nạp."""
    return kn.nap(plan_id, env)


@mcp.tool()
def pull_trang_thai(plan_id: str) -> dict:
    """Trạng thái một kế hoạch: đã kéo/nạp test/nạp prod chưa."""
    return kn.trang_thai(plan_id)


if __name__ == "__main__":
    mcp.run()

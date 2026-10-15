# -*- coding: utf-8 -*-
"""MCP server `dashboard_phieu` — kho phiếu lỗi + giao việc cho Agent tổng `quan_ly` (15/10/2026).

Chạy:  python scripts/serve_http.py phieu 8815   (HTTP/SSE cho OpenClaw gateway trong container)

Không chạm DB: mọi đọc/ghi phiếu qua API nội bộ tc-admin-api `/internal/phieu/...` (PHIEU_API_URL +
PHIEU_INTERNAL_TOKEN). Máy trạng thái ở API — tool chỉ đề nghị. Việc giao cho bot con chạy NỀN trong tiến trình
này (semaphore 3, một phiếu một việc), tool trả về ngay. Chi tiết: servers/common/phieu.py.
"""
from mcp.server.fastmcp import FastMCP

from .common import phieu as ph

mcp = FastMCP("dashboard_phieu")
ph.don_viec_treo()


@mcp.tool()
def phieu_tao_tu_log(van_ban: str, channel_id: int = None) -> dict:
    """ƯU TIÊN DÙNG khi người dùng dán log lỗi từ Excel. Nhận NGUYÊN VĂN khối dán (có/không dòng tiêu đề), tách
    TẤT ĐỊNH từng dòng -> mỗi dòng một phiếu (suy mã màn, chuẩn hoá kỳ/số/dòng mapping; dòng cùng màn+chỉ tiêu+kỳ
    -> TRUNG; 'Giống lỗi N' -> liên kết `lien_quan`, KHÔNG phải trùng). Trả `bang_markdown` để đăng cho người
    và `lo` (mã lô dán) dùng cho phieu_cot_it_fix. Không tự tách log bằng tay rồi gọi phieu_tao từng dòng.
    `channel_id` bỏ trống = kênh "Xử lý lỗi dashboard" (env PHIEU_CHANNEL_ID)."""
    return ph.tao_tu_log(van_ban, channel_id)


@mcp.tool()
def phieu_tao(nguon_log: dict, channel_id: int = None) -> dict:
    """Tạo MỘT phiếu từ `nguon_log` (schema v1: nguoi_bao, man, chi_tieu, ky bắt buộc; số là {gia_tri, don_vi}).
    Chỉ dùng khi lỗi không đến từ khối log Excel (vd người dùng kể bằng lời). API tự đăng tin gốc của luồng phiếu."""
    return ph.api_tao(nguon_log, channel_id)


@mcp.tool()
def phieu_doc(id: str) -> dict:
    """Đọc một phiếu (PL-YYYYMMDD-NNN): trạng thái, version, mọi khối nguon_log/chan_doan/quyet_dinh/sua/
    phat_hanh/nghiem_thu."""
    return ph.api_doc(id)


@mcp.tool()
def phieu_liet_ke(trang_thai: str = None, channel_id: int = None) -> dict:
    """Danh sách phiếu (lọc theo trạng thái/kênh). Muốn một bảng gọn để đăng cho người -> phieu_bang_tong."""
    ds = ph.api_liet_ke(trang_thai, channel_id)
    return {"so_phieu": len(ds), "phieu": [{"id": p["id"], "trang_thai": p["trang_thai"],
                                            "man": (p["data"].get("nguon_log") or {}).get("man"),
                                            "chi_tieu": (p["data"].get("nguon_log") or {}).get("chi_tieu"),
                                            "ky": (p["data"].get("nguon_log") or {}).get("ky")} for p in ds]}


@mcp.tool()
def phieu_bang_tong(channel_id: int = None, trang_thai: str = None) -> dict:
    """BẢNG TỔNG mọi phiếu (mã · màn · chỉ tiêu · kỳ · trạng thái · việc đang chạy · chờ ai) dạng markdown — đăng
    MỘT tin tổng hợp thay vì đọc từng phiếu."""
    return ph.bang_tong(channel_id, trang_thai)


@mcp.tool()
def phieu_chuyen(id: str, den: str, ghi_chu: str = "") -> dict:
    """Đề nghị chuyển trạng thái phiếu (vai quan_ly). API kiểm bước hợp lệ và từ chối bước sai — KHÔNG thử vòng.
    Các cổng người (DA_DUYET, HUY, DONG từ KHONG_PHAI_LOI) KHÔNG đi qua tool này: người gõ lệnh trong kênh."""
    return ph.api_chuyen(id, den, ghi_chu)


@mcp.tool()
def phieu_dang_tin(id: str, noi_dung: str) -> dict:
    """Đăng một tin TRẢ LỜI vào luồng của phiếu (bằng bot ai_quan_ly_bot)."""
    return ph.api_dang_tin(id, noi_dung)


@mcp.tool()
def phieu_giao_viec(vai: str, ids: list[str], chi_dan: str = "") -> dict:
    """Giao việc cho bot con, CHẠY NỀN, trả về NGAY (không chờ). vai ∈ tim_loi (phiếu MOI/MO_LAI/DANG_CHAN_DOAN),
    nghiem_thu (CHO_NGHIEM_THU), check, mapping, sua_loi (CHỈ khi DA_DUYET — ghi việc chờ IT, không chạy gì).
    Tối đa 3 việc chạy cùng lúc toàn hệ thống (phần còn lại xếp hàng), một phiếu một việc một lúc, 6 lần giao/phiếu.
    Kết quả bot con được kiểm schema, ghi vào phiếu, chuyển trạng thái và đăng tóm tắt vào luồng phiếu tự động."""
    return ph.giao_viec(vai, ids, chi_dan)


@mcp.tool()
def phieu_giao_hang_loat(vai: str, loc_trang_thai: str, chi_dan: str = "", channel_id: int = None) -> dict:
    """Giao cùng một vai cho MỌI phiếu đang ở `loc_trang_thai` (vd tất cả MOI -> tim_loi; CHO_NGHIEM_THU ->
    nghiem_thu). Giữ nguyên semaphore 3 và trần 6 lần giao/phiếu."""
    return ph.giao_hang_loat(vai, loc_trang_thai, chi_dan, channel_id)


@mcp.tool()
def phieu_viec(viec_id: str) -> dict:
    """Trạng thái một việc nền: cho / dang_chay / xong / loi / gian_doan, lần thử, kết quả (khối đã ghi, trạng
    thái đã chuyển, tóm tắt) hoặc lỗi."""
    return ph.doc_viec(viec_id)


@mcp.tool()
def phieu_cot_it_fix(channel_id: int = None, ids: list[str] = None, lo: str = None) -> dict:
    """Sinh nội dung 2 cột 'IT fix lỗi' · 'Trạng thái' cho từng dòng log gốc (đúng thứ tự dòng đã dán) dạng TSV
    để người dùng dán ngược vào Excel. Chỉ nêu channel_id -> lô dán gần nhất của kênh."""
    return ph.cot_it_fix(channel_id, ids, lo)


if __name__ == "__main__":
    mcp.run()

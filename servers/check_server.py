# -*- coding: utf-8 -*-
"""MCP server `dashboard_check` — bot kiểm số liệu (14/10/2026).

Chạy:  python -m servers.check_server           (stdio)
       python scripts/serve_http.py check 8812   (HTTP/SSE cho OpenClaw gateway .253)

Mọi tool đều TẤT ĐỊNH và CHỈ ĐỌC (xem servers/common/so_lieu_check.py). KHÔNG có tool nhận SQL tự do:
bot không tự tính số, chỉ gọi tool rồi diễn giải kết quả có cấu trúc.
"""
from mcp.server.fastmcp import FastMCP

from .common import man_hien_thi as mh
from .common import so_lieu_check as sc
from .common import man_fe as mf
from .common import ton_kho_xe_check as tk

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


@mcp.tool()
def check_file_ve_luc(ten_file: str, env: str = "prod") -> dict:
    """Một file nguồn (theo TÊN, vd B.2.TC.OO.D.20261012.CDPSCongnotheoRO.xlsx): sinh ở máy nguồn lúc nào,
    về 252 lần đầu/lần cuối lúc nào (giờ VN), đã vào DB những report_type/ngày nào, cron nhắc tới ở đâu.
    Dùng trả lời 'file này kéo từ thời điểm nào'."""
    return sc.file_ve_luc(ten_file, env)


@mcp.tool()
def check_so_hien_thi(man: str, tu_khoa: str, period: str = None, env: str = "prod", grain: str = "month",
                      tu_ngay: str = None, den_ngay: str = None, cong_ty: str = None, khoi: str = None) -> dict:
    """SỐ DASHBOARD ĐANG HIỂN THỊ của một chỉ tiêu: gọi đúng builder của màn rồi tìm mục khớp `tu_khoa` (nhãn thẻ như
    'Giá trị tồn kho', mã CT33, hoặc một phần tên chỉ tiêu). `man` nhận MÃ (vhkd4) HOẶC TÊN người dùng nói ('Quản lý tồn
    kho SRVF', 'phải thu XDV') — KHÔNG hỏi lại người dùng mã màn. Kết quả `khop_theo_nhan_the` khớp đúng chữ trên thẻ; không
    khớp gì thì `nhan_co_tren_man` liệt kê mọi thẻ kèm giá trị để tự chọn. `ngu_canh_cua_so` cho ngày chốt/nguồn. Trả giá trị hiển thị (thực hiện/kế hoạch/%…) và 'nguon' do
    dashboard tự khai. `period` YYYY-MM; chế độ ngày: grain='day' + tu_ngay/den_ngay (YYYY-MM-DD). Dùng khi hỏi
    'dashboard đang hiện bao nhiêu / số này có khớp file không' — rồi đối chiếu bằng tool check_* khác."""
    return mh.so_hien_thi(man, tu_khoa, period, env, grain, tu_ngay, den_ngay, cong_ty, khoi)


@mcp.tool()
def check_danh_muc_man(tu_khoa: str = None) -> dict:
    """DANH MỤC MÀN dashboard (mã màn · tên màn · nhóm · khối), đọc từ FE prod. Có `tu_khoa` (vd 'quản lý tồn kho srvf',
    'phải thu xdv') thì trả các màn khớp nhất. Dùng khi người dùng gọi màn bằng tên — đừng bắt họ đọc mã màn."""
    return {"man": mf.tim_man(tu_khoa) if tu_khoa else mf.danh_muc_man()}


@mcp.tool()
def check_ton_kho_xe(ngay: str, env: str = "prod") -> dict:
    """ĐỐI CHIẾU 'Giá trị tồn kho'/'Tồn kho hóa đơn' màn vhkd4 (Showroom/SRVF) với file nguồn KD23 của ngày `ngay`
    (YYYY-MM-DD): dòng Tổng cộng của file · xe đang tồn trong file · chân đã xuất mang dư âm · DB · số dashboard,
    kèm chênh lệch và giải thích. Dùng cho mọi câu 'giá trị/số xe tồn kho SR có đúng không, sao khác file'."""
    return tk.ton_kho_xe(ngay, env)


if __name__ == "__main__":
    mcp.run()

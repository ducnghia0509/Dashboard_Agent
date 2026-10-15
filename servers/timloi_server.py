# -*- coding: utf-8 -*-
"""MCP server `dashboard_timloi` — bot TÌM LỖI + NGHIỆM THU, CHỈ ĐỌC (15/10/2026, hợp đồng đợt 2+3 mục C).

Chạy:  python -m servers.timloi_server            (stdio)
       python scripts/serve_http.py timloi 8816    (HTTP/SSE cho OpenClaw gateway, unit deploy/dashboard-mcp-timloi.service)

Mọi tool TẤT ĐỊNH và CHỈ ĐỌC (xem servers/common/tim_loi.py, tim_loi_excel.py): DB mở default_transaction_read_only,
file chỉ trong received_reports / Tài liệu (chặn '..' và symlink), không tool nào nhận SQL hay lệnh tự do.
Thứ tự nên gọi (SKILL tim_loi): 4 tool GỘP trước (tl_ho_so_man → tl_truy_vet → tl_kiem_nhanh_nap → tl_so_qua_ky), tool lẻ
khi cần. Dữ liệu builder/hồ sơ có cache — có thể cũ tới 5 phút; truyền bo_cache=true để đo lại.
"""
from mcp.server.fastmcp import FastMCP

from .common import tim_loi as tl
from .common import tim_loi_excel as tx

mcp = FastMCP("dashboard_timloi")


# ───────────────────────── 4 tool GỘP BƯỚC (gọi trước) ─────────────────────────
@mcp.tool()
def tl_ho_so_man(man: str, tu_khoa: str = None, bo_cache: bool = False) -> dict:
    """HỒ SƠ MỘT MÀN trong một lần gọi: hàm builder (file:dòng), report_type builder đọc (+ nguồn ngày), spec trích xuất
    của từng report_type (folder, file_glob, sheet, header, cột→trường, `bay`), 3 file mới nhất trên đĩa mỗi spec (giờ sửa,
    lúc lưu về 252, có trong available_metadata không), dòng mapping của màn (hoặc khớp `tu_khoa`), nhãn thẻ FE ↔ key
    JSON. `man` nhận mã (htxt1) hoặc tên màn. Gọi ĐẦU TIÊN cho mỗi phiếu."""
    return tl.ho_so_man(man, tu_khoa, bo_cache)


@mcp.tool()
def tl_truy_vet(man: str, key: str, ky: str, env: str = "prod", so_dung: float = None, grain: str = "month",
                khoi: str = None, bo_cache: bool = False) -> dict:
    """CHUỖI file → raw_rows → builder → key cho một chỉ tiêu: số dashboard của `key` (vd 'donNhap.tongXe.nhanSl'), tổng
    raw_rows từng report_type của màn trong kỳ + source_file + cờ hidden + file mới nhất trên đĩa. Đưa `so_dung` (số đúng
    theo mapping, CÙNG đơn vị với key) để nhận `goi_y_tang_lech` (khong|builder|nap|chua_xac_dinh) — chỉ là GỢI Ý."""
    return tl.truy_vet(man, key, ky, env, so_dung, grain, khoi, bo_cache)


@mcp.tool()
def tl_kiem_nhanh_nap(doi_tuong: str, ky: str = None, env: str = "prod") -> dict:
    """BẪY TẦNG NẠP trong một lần cho một report_type (HTXT_GIAO) hoặc folder nguồn (HUNGTHINH/baocaocapnhatvanhanhxetai):
    source_file đang có dòng (đếm, tổng, kỳ), `canh_bao[]` mã HIDDEN (file trong hidden_files), CONG_DOI (≥2 file cùng
    đơn vị phủ cùng ngày/kỳ), FILE_MOI_CHUA_NAP (file về sau lượt nạp cuối / chưa có dòng), CRON_LOI (log cron gần nhất
    báo lỗi), KHONG_CO_CRON. `ky` YYYY-MM hoặc YYYY-MM-DD, bỏ trống = mọi kỳ."""
    return tl.kiem_nhanh_nap(doi_tuong, ky, env)


@mcp.tool()
def tl_so_qua_ky(env: str, man: str, key: str, cac_ky: list, grain: str = "month", khoi: str = None,
                 cong_ty: str = None, bo_cache: bool = False) -> dict:
    """MỘT key qua NHIỀU kỳ (≤24) trong một lần → bảng {ky, env, so}. env 'prod' | 'test' | 'ca_hai' (đo cả hai DB,
    kèm `lech_prod_test`). Dùng cho bước 'kiểm các tháng khác' và cho nghiệm thu nhiều kỳ."""
    return tl.so_qua_ky(env, man, key, cac_ky, grain, khoi, cong_ty, bo_cache)


# ───────────────────────── tool LẺ ─────────────────────────
@mcp.tool()
def tl_doc_mapping(tai_lieu: str = None, file: str = None, sheet: str = None, dong: int = None, den_dong: int = None,
                   tu_khoa: str = None) -> dict:
    """Dòng mapping NGUYÊN VĂN mọi cột (cột Excel + tiêu đề + giá trị). `tai_lieu` = mã mapping_index (QTTC, VHKD, XDV,
    DUAN, HTXT, ANTAXI, XANHVP, QLTS) hoặc `file` trong Tài liệu/. Không `sheet` = sheet mới nhất đã khai. `dong`..
    `den_dong` (≤60 dòng) = số dòng Excel trong log. `tu_khoa` để tìm số dòng. Trả kèm `ban_cung_ten` (bẫy bản copy)."""
    return tl.doc_mapping(tai_lieu, file, sheet, dong, den_dong, tu_khoa)


@mcp.tool()
def tl_tim_file_nguon(don_vi: str, thu_muc: str = None, tu_khoa: str = None, gioi_han: int = 60) -> dict:
    """File nguồn đã về 252 (received_reports/<ĐƠN VỊ>/<thư mục>), mới nhất trước: giờ sửa (VN), lúc lưu về 252, người
    sửa/đường ở máy nguồn, có trong available_metadata không. Không `thu_muc` và không `tu_khoa` = liệt kê thư mục con."""
    return tl.tim_file_nguon(don_vi, thu_muc, tu_khoa, gioi_han)


@mcp.tool()
def tl_doc_excel(file: str, sheet: str = None, vung: str = None, header_dong: int = None, gioi_han: int = 50,
                 tim: str = None, tu_dong: int = None) -> dict:
    """Đọc thô một sheet (chỉ đọc). `file` tương đối received_reports (DUAN/baocaotaichinhrieng/...) hoặc Tài liệu.
    Không `sheet` = danh sách sheet. `vung` 'A1:F40' = lưới ô; `header_dong` = mỗi dòng {tiêu đề: giá trị}; `tim` = chỉ
    dòng có ô bằng/chứa chuỗi đó (vd mã số '313'). Tối đa `gioi_han` (≤500) dòng."""
    return tx.doc_excel(file, sheet, vung, header_dong, gioi_han, tim, tu_dong)


@mcp.tool()
def tl_tong_hop_excel(file: str, sheet: str, header_dong: int, cot_cong: list, loc: list = None, nhom_theo: str = None,
                      tu_dong: int = None, den_dong: int = None) -> dict:
    """TỰ TÍNH SỐ ĐÚNG TỪ FILE: Σ chính xác (Decimal, không làm tròn) các cột `cot_cong` trên dòng khớp MỌI điều kiện
    `loc` = [{cot, toan_tu, gia_tri}], toan_tu ∈ = != in thang(YYYY-MM) ngay(YYYY-MM-DD) khoang_ngay([tu,den]) chua
    khong_chua. So chuỗi: bỏ dấu cách thừa, không phân biệt hoa thường (giữ dấu). Ngày: datetime / serial Excel /
    dd/mm/yyyy gõ tay. Trả số dòng khớp, ≤40 dòng góp phần (số dòng Excel), ô ngày không đọc được, ô không phải số."""
    return tx.tong_hop_excel(file, sheet, header_dong, loc, cot_cong, nhom_theo, tu_dong, den_dong)


@mcp.tool()
def tl_raw_rows(env: str, report_type: str, tu: str, den: str, nhom_theo: str = None, loc: list = None) -> dict:
    """Dòng DB `raw_rows` của một report_type trong [tu, den] (YYYY-MM-DD): tổng (Decimal), từng source_file + cờ
    `dang_an` (hidden_files — builder bỏ), `tong_bo_file_dang_an`, theo `nhom_theo`, ≤40 dòng mẫu. Cột lọc/nhóm:
    source_file, ngay, period_month, cong_ty, khoi, cost_center, dim1..dim3, dataset_kind, payload.<khoá>."""
    return tl.raw_rows(env, report_type, tu, den, nhom_theo, loc)


@mcp.tool()
def tl_so_dashboard(env: str, man: str, ky: str, grain: str = "month", khoi: str = None, key: str = None,
                    cong_ty: str = None, tu_ngay: str = None, den_ngay: str = None, bo_cache: bool = False) -> dict:
    """SỐ DASHBOARD builder trả (đúng số FE vẽ, quyền admin) cho (màn, kỳ). `key` đường JSON ('phaiNop.cuoi',
    'kpi.sl.thucHien' — đoạn trong danh sách khớp theo ma/id/name/…) -> `gia_tri` + `so`. Không key = JSON đầy đủ (rút
    gọn khi quá lớn). grain='day': ky YYYY-MM-DD hoặc tu_ngay/den_ngay. Cache 5 phút; bo_cache=true để dựng lại."""
    return tl.so_dashboard(env, man, ky, grain, khoi, key, cong_ty, tu_ngay, den_ngay, bo_cache)


@mcp.tool()
def tl_doc_code(man: str = None, ham: str = None, env: str = "prod") -> dict:
    """Mã nguồn (chỉ đọc) builder của màn (`man`: file:dòng, docstring, report_type, thân hàm) hoặc một hàm bất kỳ
    trong app/ của tc-admin-api (`ham` = tên hàm, vd '_where'). env=prod đọc cây prod, env=test đọc cây TEST."""
    return tl.doc_code(man, ham, env)


if __name__ == "__main__":
    mcp.run()

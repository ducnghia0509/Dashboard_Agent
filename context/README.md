# Ngữ cảnh cho agent analyst — NGUỒN SỰ THẬT

**Nguồn máy đọc (single source of truth) = `~/template_trust/Template_chuan.xlsx`.**
Agent analyst đọc LIVE mỗi lần analyse qua `contract.guide()` / `template_contract_info`
(không cache). Các file khác chỉ là INPUT do người/kế toán soạn, KHÔNG được máy đọc trực tiếp.

## Sửa ngữ cảnh ở đâu
| Muốn đổi | Sửa ở | Rồi chạy |
|---|---|---|
| 50 chỉ tiêu / công thức / sheet+cột nguồn / màn FE / trạng thái dữ liệu | `knowledge/50_chi_tieu.yaml` (2026-07-09, thay `Template_chuan!00_50CHITIEU` làm nguồn nội dung) | `python scripts/gen_kpi_glossary.py` (sinh lại `kpi_glossary.json` cho QA/analyst) |
| Ngưỡng cảnh báo (nguong_vang/nguong_do/trang_thai_den) — CHƯA có trong yaml trên | `Template_chuan.xlsx`, sheet `00_50CHITIEU` (join best-effort theo tên chỉ tiêu, có thể lệch/rỗng) | `python scripts/gen_kpi_glossary.py` |
| Công ty / pháp nhân | `Template_chuan.xlsx`, sheet `MD_CONGTY` | (không cần; đọc live) |
| Khối kinh doanh | `MD_KHOIKD` | (không cần) |
| Cost center → công ty → khối | `MD_COSTCENTER` | (không cần) |
| Quy tắc chiều (HQKD theo Khối, Dòng tiền theo Pháp nhân) | `bo_sung_nguoi_dung.txt` (file này, nạp vào prompt) | (không cần) |

## sources/ — file INPUT gốc (tham chiếu lịch sử, KHÔNG máy đọc)
- `Huong_dan_cach_lay_va_tinh_toan.xlsx` — bản guideline 50 chỉ tiêu do nghiệp vụ soạn (2026-07).
  Nội dung ĐÃ được hợp nhất vào `Template_chuan!00_50CHITIEU` (bản đã map sheet/cột — đầy đủ hơn).
- `Danh_Muc_Ma_he_thong.xlsx` — danh mục mã (khối/pháp nhân/cost center/quy tắc mã tên file).
  Nội dung ĐÃ có trong `Template_chuan!MD_CONGTY/MD_KHOIKD/MD_COSTCENTER`.

## backup_20260706/ — bản lưu trước khi hợp nhất
- `Template_chuan.xlsx.bak`, `kpi_glossary.json.bak`
- `guildline.xlsx.ARCHIVED` — guideline rời CŨ (sai chính tả tên), đã NGỪNG dùng. gen_kpi_glossary.py
  trước đây trỏ nhầm `guideline.xlsx` (thiếu chữ) → script gãy, glossary bị stale. Nay sinh từ Template.

## KHÔNG phải ngữ cảnh analyst (để riêng, đừng gộp)
- `servers/common/canonical.py` — router loại báo cáo tất định (KQKD/CDKT/TK131…). Hardcode có chủ đích.
- `display_contract.json` — bản đồ field ↔ màn hình FE (phân tích thiếu nguồn), mối quan tâm khác.

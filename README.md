# DashBoard_Agent

OpenClaw (Claude Code) multi-agent cho `DashBoard_AI`: 1 orchestrator + 3 subagent
(`analyst`, `execute`, `qa`), 2 MCP server Python. Không dùng RAG/vector/embeddings —
mọi tra cứu là deterministic (SQL thật + JSON glossary/discovery). Xem `plan.md` gốc
trong `DashBoard_AI/plan.md` để hiểu bối cảnh đầy đủ; file này chỉ nói cách chạy.

## 2 luồng

```
Luồng 1 (ingest):  file xlsx lạ -> analyst (template_analyze, đối chiếu
                   display_contract.json + kpi_glossary.json) -> [human approve]
                   -> execute (dry-run -> import_execute) -> raw_rows

Luồng 2 (qa):      câu hỏi tiếng Việt -> qa (sql_query / glossary_lookup /
                   discovery_search / source_inspect) -> trả lời + bảng + nguồn
```

## Kiến trúc 2 tầng: OpenClaw (agent/tool) + 9Router (model routing)

OpenClaw (chính là hệ thống orchestrator + subagent + MCP trong thư mục này) và
9Router giải quyết hai tầng khác nhau, KHÔNG chồng lấn:

- **OpenClaw** (tầng agent): quản lý orchestrator/subagent, tool permission, MCP,
  memory, session. Đây là những gì `agents/*/SKILL.md` + `servers/*.py` mô tả.
- **9Router** (tầng model): quản lý model, routing, fallback, quota, cung cấp 1
  endpoint OpenAI-compatible duy nhất để đổi model mà không phải sửa agent nào.

```
                    User
                      │
                 OpenClaw
                (Orchestrator)
                      │
     ┌────────────────┼─────────────────┐
     │                │                 │
 analyst         execute             qa
(subagent)      (subagent)        (subagent)
     │                │                 │
     └────────────── MCP ───────────────┘
                      │
         dashboard_ingest / dashboard_qa
                      │
                DashBoard_AI
                      │
                  PostgreSQL
```

Toàn bộ model đi qua 9Router (chưa triển khai — đây là thiết kế cho tương lai, chưa
có server nào chạy thật ở `localhost:20128`):

```
OpenClaw
     │
OpenAI-compatible API
     │
localhost:20128
     │
9Router
     │
──────────────────────────────────────
Claude · GPT-5 · Gemini · DeepSeek · MiniMax · Qwen · ...
```

Lợi ích: mỗi subagent có thể dùng 1 model khác nhau tuỳ độ khó việc nó làm (orchestrator
cần model mạnh để điều phối/quyết định approve; execute gần như chỉ gọi tool nên dùng
model rẻ) — mà tất cả vẫn gọi chung 1 endpoint `http://localhost:20128/v1`, đổi model
chỉ cần sửa cấu hình 9Router, không phải sửa `SKILL.md`/code agent nào. 9Router cũng
cho fallback nếu 1 provider hết quota/lỗi.

### Model + quyền tool theo từng subagent

| Subagent | Model mặc định (qua 9Router) | Tool MCP được phép gọi |
|---|---|---|
| `orchestrator` | `claude-sonnet` | không gọi tool trực tiếp — chỉ điều phối + spawn subagent |
| `analyst` | `minimax` | `discover_files`, `template_analyze`, `sheet_profile`, `glossary_lookup`, `discovery_search`, `report_spec_search` |
| `execute` | `qwen` | `import_execute`, `generic_import_execute` (luôn `dry_run=true` trước, rồi mới `dry_run=false`) |
| `qa` | `gpt-5-mini` | `sql_query`, `glossary_lookup`, `discovery_search`, `report_spec_search`, `source_inspect`, `schema_describe` |

Model cụ thể ghi trong frontmatter mỗi `SKILL.md` (`model: ...`) — đổi model cho 1
agent chỉ cần sửa dòng đó + cấu hình route tương ứng bên 9Router, không phải sửa
logic prompt. Danh sách tool ở trên là quyền TỐI ĐA cho từng subagent (giữ đúng
triết lý giới hạn tool per-agent của OpenClaw) — `import_plan_validate` và
`discovery_record_tool` trong `ingest_server.py` vẫn tồn tại làm tool phụ/thủ công
nhưng không nằm trong quyền mặc định của agent nào ở trên (validate giờ nằm trong
bước `dry_run=true` của `import_execute`; `template_analyze` đã tự ghi discovery
memory nên không cần agent gọi `discovery_record_tool` thủ công nữa).

## Sheet lạ không khớp report_type cố định (`sheet_profile` + `generic_import_execute`)

Nhiều file thật (vd báo cáo tài chính riêng dạng `B01/B02/B03-DN` với sheet `131`/`331`/
`Biểu khấu hao`/`KQKD`/`CDKT`) không khớp 9 report_type cố định (`FIELD_DEFS`) lẫn tên sheet
cố định của `importer_month.py` (`KQKD_HỢP NHẤT THÁNG`/`CĐKT_HỢP NHẤT`) — `template_analyze`
sẽ trả `report_type=None`. Với ca này:

1. `analyst` gọi `sheet_profile` để tự khám phá cấu trúc sheet (xem 10 dòng đầu theo chiều
   dọc + 8 cột đầu theo chiều ngang để phát hiện header 1-2 dòng và layout thực thể-theo-cột),
   rồi `glossary_lookup`/`report_spec_search` để khớp với `guideline.xlsx`, tự dựng 1
   `SheetMapping` khai báo (không sinh code) — xem `servers/common/models.py`.
2. `execute` gọi `generic_import_execute` — diễn giải `SheetMapping` theo 1 thuật toán cố định
   ("melt" bảng rộng thành `raw_rows` dạng dài, tổng quát hoá pattern `_parse_cdkt` trong
   `importer_month.py`), ghi report_type mới với tiền tố **`GEN_`** (vd `GEN_SO131_PHAITHU`) để
   tách biệt hẳn 10 mã cố định. Tool này LUÔN trả `sample_mapped_rows` để người xem trước khi approve.
3. Mapping học được lưu vào catalog tích luỹ `memory/report_specs/<fingerprint>.json` (gồm cả
   `display_spec` — định dạng CHỈ để FE hiển thị sau này, không dùng khi ghi `raw_rows`) — lần
   sau gặp sheet cùng fingerprint không phải suy luận lại; `qa` cũng đọc catalog này qua
   `report_spec_search` làm ngữ cảnh khi giải thích số liệu `GEN_*` (không cần định dạng riêng).

Đã verify thật với file `B.7.AAG.TCKT.M.202601.Baocaotaichinhrieng.xlsx` (không có trong repo,
do user cung cấp riêng): `sheet_profile` phát hiện đúng header 2 dòng của sheet `131` (dòng
"Đầu kỳ/Phát sinh/Cuối kỳ" rồi "Nợ/Có"); `generic_import_execute(dry_run=True)` melt đúng 93
dòng khớp thủ công với dữ liệu thô (khách hàng `02010001`: Đầu kỳ Nợ=359.409.087, PS Nợ=
401.964.000, PS Có=264.630.251, Cuối kỳ Nợ=496.742.836).

### `discover_files` + `canonical_kind` — mở rộng cho nhiều công ty/nhiều file

Khi có nhiều công ty, mỗi công ty tự đặt tên sheet/tên file khác nhau cho CÙNG 1 loại báo cáo
kế toán (vd công ty A đặt sheet `131`, công ty B đặt `Cong no phai thu`) — 2 cải tiến sau giúp
`analyst` không phải suy luận lại từ đầu cho mỗi công ty, mà KHÔNG cần xây catalog SQL/pandas
query engine riêng (cân nhắc rồi quyết định KHÔNG làm — xem lý do bên dưới):

- `discover_files(root_dir, pattern)` (tool mới, `dashboard_ingest`) — quét `INPUT_DIR` (hoặc
  thư mục con), trả danh sách file kèm `company_guess` (từ tên thư mục con, hoặc regex mã công
  ty trong tên file vd `B.7.AAG...` → `AAG`) và `canonical_kind_guess` (từ tên file). RẺ — không
  mở nội dung workbook.
- `canonical_kind` (`servers/common/canonical.py`) — 1 dict alias NHỎ, cố định trong code (KHÔNG
  tách thành `aliases.json` riêng để tránh nhân bản nguồn sự thật với `kpi_glossary.json`/
  `FIELD_DEFS` đã có), đoán loại báo cáo chuẩn hoá (`TK131`/`TK331`/`TSCD`/`KQKD`/`CDKT`/`CDPS`/
  `LCTT`) từ tên sheet/tên file. `sheet_profile(sheet=None)` tự gắn `canonical_kind_guess` cho
  mỗi sheet; `SheetMapping`/`report_specs` catalog lưu thêm field `canonical_kind` +
  `company` để `report_spec_search(canonical_kind=...)` tìm được mapping đã học ở CÔNG TY KHÁC.
- **Cân nhắc nhưng KHÔNG làm**: 1 query engine song song (pandas/SQL-trên-dataframe) cho dữ liệu
  chưa import, và chuyển hẳn metadata sang SQL catalog (4 bảng) ngay bây giờ. Lý do: vi phạm
  nguyên tắc "1 nguồn sự thật qua `raw_rows`+SQL" của dự án (dễ ra 2 số khác nhau cho cùng câu
  hỏi), và over-engineer ở quy mô hiện tại (1 file thật + 6 file mẫu). `sql/create_discovery_table.sql`
  đã viết sẵn làm tiền lệ — khi thật sự có hàng trăm/nghìn file thì chuyển catalog JSON hiện tại
  sang SQL theo đúng mẫu đó.

## Cài đặt

```bash
cd DashBoard_Agent
python -m venv .venv
.venv/Scripts/activate        # Windows; source .venv/bin/activate trên Linux/Mac
pip install -r requirements.txt
cp .env.example .env          # rồi điền DATABASE_URL/DATABASE_URL_RO thật khi sẵn sàng
```

`servers/common/be_bridge.py` tự `sys.path.insert` vào `BACKEND_PATH` (mặc định
`../DashBoard_AI/backend`) và import lại nguyên `app.schemas/importer/importer_month/
importer_ledger/db/repo/master/metrics` — không viết lại logic nghiệp vụ.

## Sinh kpi_glossary.json (chạy 1 lần, hoặc mỗi khi guideline.xlsx đổi)

```bash
python scripts/gen_kpi_glossary.py
```

Đọc `../DashBoard_AI/guideline.xlsx` (50 chỉ số quản trị), ghi `kpi_glossary.json`.
Mỗi record có `needs_followup=true` nếu nguồn dữ liệu chưa được hệ thống hiện tại
parse (vd TK131/TK331/Sheet TSCĐ/CĐSPS chi tiết/aging) — `qa`/`analyst` dùng cờ này
để tránh bịa số cho các chỉ số chưa có nguồn.

## Tự-verify không cần mở Claude Code (không cần DB thật)

```bash
python scripts/smoke_test.py       # guardrails + template_analyze + discovery + glossary
python eval/template/score.py      # chấm analyst trên 6 file mẫu Data_test_dashboard/
python eval/qa_golden/score.py     # chấm qa trên 20 câu hỏi (phần sql tự skip nếu thiếu .env)
```

Lưu ý: `template_analyze` tái dùng `importer.prepare()` của DashBoard_AI, hàm này có
1 bước `SELECT` read-only trên bảng `mapping_profiles` (Postgres thật, DATABASE_URL đọc
từ `DashBoard_AI/.env`) để tra cache mapping — không ghi, không DDL, đúng thiết kế cache
sẵn có của hệ thống.

## Bật luồng thật (tự làm khi sẵn sàng — KHÔNG tự động chạy trong quá trình build)

1. Chạy `sql/create_ro_role.sql` thủ công trên Postgres thật (đổi password trước).
2. Điền `DATABASE_URL` (ghi) và `DATABASE_URL_RO` (role vừa tạo) vào `.env`.
3. Đăng ký MCP server trong Claude Code: trỏ file cấu hình MCP tới `mcp_config.json`
   (2 server `dashboard_ingest` và `dashboard_qa`, chạy bằng
   `python -m servers.ingest_server` / `qa_server`, `cwd` = thư mục `DashBoard_Agent`).
4. Mở Claude Code tại đây, dùng skill `agents/orchestrator/SKILL.md` làm điểm vào.
5. Thử luồng 1: đưa 1 file trong `Data_test_dashboard/` cho orchestrator, duyệt
   ImportPlan, xác nhận nạp — `execute` sẽ dry-run rồi mới ghi thật (idempotent, xem
   `memory/imports_ledger.json`).
6. Thử luồng 2: hỏi qa 1 câu trong `eval/qa_golden/questions.json` (loại `sql`) và so
   với `/metrics/...` tương ứng trên DashBoard_AI.

## Cấu trúc

```
agents/{orchestrator,analyst,execute,qa}/SKILL.md   prompt cho từng subagent
servers/common/*        be_bridge, models, guardrails, db_ro/db_rw, introspect, memory
servers/ingest_server.py / qa_server.py             2 MCP server (FastMCP, stdio)
scripts/gen_kpi_glossary.py / smoke_test.py
display_contract.json   12 màn hình FE: field thật (wired_fields) vs hard-code (static_fields)
kpi_glossary.json        50 chỉ số quản trị từ guideline.xlsx (sinh bởi script trên)
sql/                     create_ro_role.sql (chạy thủ công), create_discovery_table.sql (dự phòng)
memory/discoveries/      discovery memory dạng JSON (không vector)
memory/report_specs/     catalog SheetMapping đã học cho sheet lạ (report_type GEN_*)
eval/                    eval/template (6 file mẫu) + eval/qa_golden (20 câu)
```

## Rủi ro / lưu ý đã biết

- Nhiều chỉ số trong `kpi_glossary.json` (`needs_followup=true`) cần nguồn dữ liệu mà
  `DashBoard_AI` hiện chưa parse (TK131, TK331, Sheet TSCĐ, CĐSPS chi tiết...) — đây là
  giới hạn của dữ liệu đầu vào hiện có, KHÔNG phải lỗi của agent. `qa` phải nói rõ
  "chưa có nguồn tự động" thay vì suy diễn số.
- `import_execute` idempotent bằng `memory/imports_ledger.json` (khoá theo
  dataset_kind+report_type+period+fingerprint) — đây là cơ chế riêng của
  DashBoard_Agent, không sửa schema Postgres của DashBoard_AI.
- Không sửa gì trong `DashBoard_AI/` — chỉ đọc (`be_bridge`) để tái dùng logic.

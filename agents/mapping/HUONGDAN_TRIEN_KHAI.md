# Triển khai Bot AI So khớp mapping (`ai_mapping_bot`, agent `mapping`)

Chạy trên OpenClaw LOCAL của 252 như bot `check`/`keo_file`. MCP chạy trên host, gắn vào docker bridge.

1. **MCP server (sudo — bạn chạy).** Unit giống `dashboard-mcp-check.service`: đổi tên `dashboard-mcp-mapping`,
   lệnh `scripts/serve_http.py mapping 8814`, bind 172.17.0.1. `systemctl enable --now`. Kiểm: `curl -sI http://172.17.0.1:8814/sse` → 200.
2. **Đăng ký MCP trong container:** `docker exec openclaw-openclaw-gateway-1 openclaw mcp add dashboard_mapping --url http://172.17.0.1:8814/sse --transport sse`
3. **Tạo agent:** `openclaw agents add mapping`; chép `agents/mapping/SKILL.md` → `workspace/skills/so_khop_mapping/SKILL.md`
   và `workspace/{IDENTITY,SOUL,AGENTS}.md`.
4. **tc-admin-api:** thêm `ai_mapping_bot=mapping` vào `OPENCLAW_DOCKER_BOTS` trong .env, chạy
   `python tools/tao_bot_ai_check.py --bot mapping --env test --write` (rồi prod), restart API, gắn bot vào MỘT kênh riêng.

Nghiệm thu: (a) "mapping đã đủ cho các màn chưa?" (b) "chỉ tiêu tỷ lệ hoàn thành doanh thu ở vhkd0 mapping với code có khớp không"
(c) "màn at1 đọc nguồn nào". Đối chiếu `scripts/test_mapping_check.py`.
Lưu ý: cache logic ở `memory/mapping_bot/logic.json` — xoá file khi code builder đổi.

# Triển khai Bot AI Kéo file (`ai_keo_file_bot` → agent `keo_file`)

Chạy trên OpenClaw LOCAL của 252 (như bot kiểm số liệu). Cần 2 PR đã merge: Dashboard_Agent (`feat/bot-keo-file`) và
tc-admin-api (`feat/bot-keo-file-xac-nhan`).

1. **MCP server `dashboard_pull` (sudo):** unit y hệt `dashboard-mcp-check`, đổi tên + cổng 8813:
   `ExecStart=.../.venv/bin/python scripts/serve_http.py pull 8813`, `Environment=MCP_BIND_HOST=172.17.0.1`.
2. **Đăng ký MCP + tạo agent trong container:**
   ```bash
   docker exec openclaw-openclaw-gateway-1 openclaw mcp add dashboard_pull --url http://host.docker.internal:8813/sse --transport sse
   R=/home/itadmin/AI_Dashboard_QT/Dashboard_Agent/agents/keo_file; W=/home/node/.openclaw/agents/keo_file/workspace
   docker exec openclaw-openclaw-gateway-1 sh -c "mkdir -p $W/skills/keo_nap_file && cp $R/SKILL.md $W/skills/keo_nap_file/SKILL.md && cp $R/workspace/IDENTITY.md $R/workspace/SOUL.md $R/workspace/AGENTS.md $W/"
   docker exec openclaw-openclaw-gateway-1 openclaw agents add keo_file --workspace $W --model 9router/my1 --non-interactive
   docker restart openclaw-openclaw-gateway-1
   ```
3. **API prod:** thêm bot vào định tuyến local (giữ bot kiểm số liệu): trong `.env`
   `OPENCLAW_DOCKER_BOTS=ai_check_bot=check,ai_keo_file_bot=keo_file` rồi restart `tc-admin-api` và `tc-bot-worker`.
4. **Đăng ký bot trong dashboard** (tc-admin-api): `python tools/tao_bot_ai_check.py --bot keo_file --env test --write` rồi `--env prod`.
5. **Gắn bot vào một kênh riêng** chỉ gồm người được phép (mỗi kênh một bot; ai ở trong kênh đều gọi được bot).

Nhật ký thao tác ghi: `Dashboard_Agent/memory/pull_bot/audit.jsonl`. Xác nhận PROD: `memory/pull_bot/xacnhan/<MÃ>.json`
(do API ghi, dùng một lần, hết hạn 30 phút).

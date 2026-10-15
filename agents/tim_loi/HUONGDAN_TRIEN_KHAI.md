# Triển khai MCP `dashboard_timloi` + agent `tim_loi`, `nghiem_thu` (15/10/2026)

Code (tool chỉ đọc) nằm ở `servers/timloi_server.py`, `servers/common/tim_loi.py`, `servers/common/tim_loi_excel.py`,
`scripts/timloi_so_dump.py`. Thứ tự tổng thể (migration, .env, tạo kênh…) xem `deploy/agent_tong_HUONGDAN.md` (mục D hợp đồng).

1. **Kiểm trước (AI/IT, không cần sudo):** `.venv/bin/python scripts/test_timloi_eval.py` → `… PASS, 0 FAIL` (đọc PROD
   read-only; `--env test` để chạy trên 5435).
2. **MCP server trên 252 (cần sudo — người chạy):**
   `sudo cp deploy/dashboard-mcp-timloi.service /etc/systemd/system/ && sudo systemctl daemon-reload && sudo systemctl enable --now dashboard-mcp-timloi`
   Kiểm: `curl -s -m 3 -o /dev/null -w '%{http_code}\n' http://172.17.0.1:8816/sse` (200; lệnh treo vì SSE mở luồng — `-m 3` cắt).
3. **Đăng ký với gateway:** `openclaw mcp add dashboard_timloi --url http://host.docker.internal:8816/sse --transport sse`
4. **Agent:** tạo `tim_loi`, `nghiem_thu`; chép `agents/<agent>/SKILL.md` → `workspace/skills/<agent>/SKILL.md` và
   `agents/<agent>/workspace/{AGENTS,IDENTITY,SOUL}.md` → `workspace/`. Sau này sửa SKILL thì
   `OPENCLAW_AGENTS="tim_loi nghiem_thu" scripts/sync_skills.sh`.
5. **Giới hạn tool theo agent** (`deploy/openclaw_tools_theo_agent.json` + `scripts/ap_tool_theo_agent.py`):
   - `tim_loi`: allow `read`, `dashboard_timloi__*`, `dashboard_mapping__*`, `dashboard_check__*`
   - `nghiem_thu`: allow `read`, `dashboard_timloi__tl_so_dashboard`, `dashboard_check__check_so_hien_thi`
   - cả hai: `profile: coding`, deny `group:runtime`, `write`, `edit`, `apply_patch`; `fs.workspaceOnly: true`.

Lưu ý vận hành: số builder cache 5 phút trong tiến trình MCP (`bo_cache=true` để đo lại); `tl_so_dashboard` chạy builder
TRÊN ĐĨA cây prod (`/home/itadmin/apps/tc-console/tc-admin-api`) — merge+pull mà chưa restart API thì số tool ra có thể
khác số màn hình đang hiện.

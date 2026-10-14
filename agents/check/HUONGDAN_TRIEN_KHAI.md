# Triển khai Bot AI Kiểm số liệu (`ai_check_bot`)

Phần code (tool tất định) đã nằm ở 252. Còn 4 bước hạ tầng, mỗi bước ghi rõ AI làm hay bạn làm.

1. **MCP server trên 252 (cần sudo — bạn chạy).** Chạy `scripts/serve_http.py check 8812` như `qa` (8811).
   Cách làm giống unit `dashboard-mcp-qa.service`: sao chép unit, đổi tên + cổng, `systemctl enable --now`.
   Kiểm nhanh: `curl -sI http://127.0.0.1:8812/sse` trả 200.
2. **Đăng ký MCP trên gateway .253 (bạn/IT .253):**
   `openclaw mcp add dashboard_check --url http://host.docker.internal:8812/sse --transport sse`
3. **Tạo agent trên .253:** `openclaw agents add ai_check_bot`; chép `agents/check/SKILL.md` vào
   `workspace/skills/check_so_lieu/SKILL.md` và `workspace/{IDENTITY,SOUL,AGENTS}.md`; `gateway restart`.
   Nhớ: định tuyến cần cả header `x-openclaw-agent-id` lẫn session key `agent:ai_check_bot:<...>`.
4. **Gắn bot vào kênh chat trên 252:** thêm `ai_assistants` + `chat_actors` (`code = ai_check_bot`) như
   Bot AI IT (script idempotent, mặc định dry-run), rồi gắn vào MỘT kênh (một kênh một bot).

Nghiệm thu: hỏi 3 câu mẫu — (a) "so sánh BC tháng 9 và tổng BC ngày của An Taxi",
(b) "phải thu XDV ngày 12/10 lấy ở file nào, vì sao khác tổng file", (c) "HO sao chưa có số T10" —
đối chiếu với `scripts/test_so_lieu_check.py` (16 ca) trên DB test.

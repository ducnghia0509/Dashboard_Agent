# Triển khai Agent tổng (`ai_quan_ly_bot` → agent `quan_ly`) + bot tìm lỗi / nghiệm thu

Hợp đồng: `/home/itadmin/AI_Dashboard_QT/wt/hopdong_dot2_agent_tong_20261015.md` (mục D). NGƯỜI chạy từng bước, đúng
thứ tự. Mỗi bước có lệnh kiểm; bước nào kiểm hỏng thì DỪNG, đừng chạy bước sau.

Cần merge trước (3 nhánh):
| repo | nhánh | nội dung |
|---|---|---|
| tc-admin-api | `feat/phieu-loi-agent-tong` | migration `bug_tickets`, `/internal/phieu/*`, lệnh người, `tools/tao_kenh_agent_tong.py` |
| Dashboard_Agent | `feat/mcp-phieu-quan-ly` | MCP `dashboard_phieu` (8815), agent `quan_ly`, hướng dẫn này, tools theo agent |
| Dashboard_Agent | `feat/mcp-timloi-chi-doc` | MCP `dashboard_timloi` (8816, chỉ đọc), agent `tim_loi` + `nghiem_thu`, `deploy/dashboard-mcp-timloi.service` |

**Merge `feat/mcp-timloi-chi-doc` TRƯỚC hoặc CÙNG LÚC với `feat/mcp-phieu-quan-ly`**: nhánh phiếu khai tools cho
`tim_loi`/`nghiem_thu` trong `deploy/openclaw_tools_theo_agent.json` và hướng dẫn này chép SKILL từ
`agents/tim_loi`, `agents/nghiem_thu` — hai thư mục đó chỉ có ở nhánh timloi. Cả hai nhánh cùng sửa
`scripts/serve_http.py` nhưng ở các dòng khác nhau (`git merge-tree` thử ngày 15/10: merge sạch, có cả `phieu` lẫn
`timloi`).

Ký hiệu: `API=/home/itadmin/apps/tc-console/tc-admin-api` (PROD), `AG=/home/itadmin/AI_Dashboard_QT/Dashboard_Agent`,
`OC="docker exec openclaw-openclaw-gateway-1"`.

## 0. Kéo code về máy chạy
```bash
git -C $API pull --ff-only
git -C $AG pull --ff-only
ls $API/migrations/*bug_tickets*.sql $AG/servers/phieu_server.py $AG/servers/timloi_server.py   # đủ 3 thứ
```

## 1. Migration (TEST trước, PROD sau)
Tên file theo nhánh API lúc viết hướng dẫn: `0093_bug_tickets.sql` (+ `0093_bug_tickets.down.sql` để hoàn tác).
Nếu khi merge số đã đổi, dùng đúng tên trong `ls` ở bước 0.
```bash
cd $API
psql "postgresql://tc:<MẬT_KHẨU>@localhost:5435/tc_dashboard" -v ON_ERROR_STOP=1 -f migrations/0093_bug_tickets.sql
psql "postgresql://tc:<MẬT_KHẨU>@localhost:5434/tc_dashboard" -v ON_ERROR_STOP=1 -f migrations/0093_bug_tickets.sql
# kiểm
psql "postgresql://tc:<MẬT_KHẨU>@localhost:5434/tc_dashboard" -c '\d bug_tickets' -c '\d bug_ticket_events'
# hoàn tác (nếu cần): ... -f migrations/0093_bug_tickets.down.sql
```

## 2. `.env`
Sinh một khoá dùng chung cho API và MCP (không dán khoá vào chat):
```bash
python3 -c "import secrets; print(secrets.token_urlsafe(32))"
```
**`$API/.env`** (API + tc-bot-worker đọc chung — `BOT_WORKER_BAT=1` nên lượt bot đi qua worker):
```
OPENCLAW_DOCKER_BOTS=ai_check_bot=check,ai_keo_file_bot=keo_file,ai_mapping_bot=mapping,ai_quan_ly_bot=quan_ly
PHIEU_INTERNAL_TOKEN=<khoá vừa sinh>
PHIEU_NGUOI_DUYET=                # để TRỐNG lúc này (= không ai duyệt được); điền ở bước 9
```
(Giữ nguyên ba bot đang có trong `OPENCLAW_DOCKER_BOTS`; chỉ nối thêm `ai_quan_ly_bot=quan_ly`.)
Tuỳ chọn: `PHIEU_REPO_DIRS` / `PHIEU_REPO_UNITS` để API tự kiểm merge/restart cho repo khác tc-admin-api (mặc định
chỉ biết tc-admin-api — repo khác thì người tự `CHUYEN`).

**`$AG/.env`** (MCP `dashboard_phieu` đọc qua `scripts/serve_http.py`):
```
PHIEU_API_URL=http://127.0.0.1:8011
PHIEU_INTERNAL_TOKEN=<cùng khoá>
PHIEU_CHANNEL_ID=                 # điền ở bước 9 (id kênh "Xử lý lỗi dashboard")
# PHIEU_AGENT_TIMEOUT=590         # giây chờ MỘT lượt bot con (mặc định 590 như OPENCLAW_TIMEOUT_SECONDS).
#                                 # Nâng lên >600 thì phải nâng cả timeoutSeconds của agent tim_loi trong openclaw.json.
```

## 3. Restart API + worker (sau migration + .env)
```bash
sudo systemctl restart tc-admin-api tc-bot-worker
systemctl is-active tc-admin-api tc-bot-worker
# endpoint nội bộ sống + token đúng (trả [] khi chưa có phiếu):
curl -s -H "X-Phieu-Token: $(grep ^PHIEU_INTERNAL_TOKEN= $API/.env | cut -d= -f2-)" http://127.0.0.1:8011/internal/phieu
curl -s -o /dev/null -w '%{http_code}\n' http://127.0.0.1:8011/internal/phieu       # không token -> 401/403/404
```

## 4. Cài 2 unit MCP (8815 phiếu, 8816 tìm lỗi)
```bash
sudo cp $AG/deploy/dashboard-mcp-phieu.service $AG/deploy/dashboard-mcp-timloi.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now dashboard-mcp-phieu dashboard-mcp-timloi
curl -s -m 2 -o /dev/null -w '%{http_code}\n' http://172.17.0.1:8815/sse    # 200 (Ctrl-C nếu treo luồng SSE)
curl -s -m 2 -o /dev/null -w '%{http_code}\n' http://172.17.0.1:8816/sse    # 200
journalctl -u dashboard-mcp-phieu -n 5 --no-pager                            # "[serve_http] phieu on 172.17.0.1:8815"
```
`dashboard-mcp-phieu` chạy user `itadmin` (nhóm docker) vì việc nền gọi `docker exec … openclaw agent`. Trạng thái
việc: `$AG/memory/phieu/viec/*.json`; việc chờ IT sửa: `$AG/memory/phieu/viec_sua/<PL-…>.json`; nhật ký
`$AG/memory/phieu/audit.jsonl`.

## 5. Đăng ký 2 MCP với gateway OpenClaw
```bash
$OC openclaw mcp add dashboard_phieu  --url http://host.docker.internal:8815/sse --transport sse
$OC openclaw mcp add dashboard_timloi --url http://host.docker.internal:8816/sse --transport sse
$OC openclaw mcp list            # thấy dashboard_phieu, dashboard_timloi (mcp add tự probe: unit bước 4 phải đang chạy)
```

## 6. Tạo 3 agent + workspace + chép SKILL
`~/.openclaw` thuộc user `node` của container → mọi ghi đi qua `docker exec -i … sh -c "cat > …"` (đường dẫn host
KHÔNG thấy được từ trong container, đừng `cp` thẳng từ `$AG`).
```bash
W=/home/node/.openclaw/agents
chep() {  # chep <agent> <tên thư mục skill> — SKILL + IDENTITY/SOUL/AGENTS của repo vào workspace
  a=$1; sk=$2; R=$AG/agents/$a
  $OC sh -c "mkdir -p $W/$a/workspace/skills/$sk"
  docker exec -i openclaw-openclaw-gateway-1 sh -c "cat > $W/$a/workspace/skills/$sk/SKILL.md" < $R/SKILL.md
  for f in IDENTITY.md SOUL.md AGENTS.md; do
    docker exec -i openclaw-openclaw-gateway-1 sh -c "cat > $W/$a/workspace/$f" < $R/workspace/$f
  done
}
chep quan_ly    quan_ly_phieu_loi
chep tim_loi    tim_loi
chep nghiem_thu nghiem_thu
# tên thư mục skill phải khớp đường dẫn mà agents/<agent>/workspace/AGENTS.md bảo đọc:
grep -h "skills/" $AG/agents/{quan_ly,tim_loi,nghiem_thu}/workspace/AGENTS.md

for a in quan_ly tim_loi nghiem_thu; do
  $OC openclaw agents add $a --workspace $W/$a/workspace --model 9router/my1 --non-interactive
done
$OC sh -c "ls $W/quan_ly/workspace/skills/*/SKILL.md $W/tim_loi/workspace/skills/*/SKILL.md $W/nghiem_thu/workspace/skills/*/SKILL.md"
```
Sửa SKILL về sau: `OPENCLAW_AGENTS="quan_ly tim_loi nghiem_thu" bash $AG/scripts/sync_skills.sh` (có restart gateway).

## 7. Giới hạn tool theo agent
Nội dung: `$AG/deploy/openclaw_tools_theo_agent.json` — `quan_ly`: read + `dashboard_phieu__*`; `tim_loi`: read +
`dashboard_timloi__*` + `dashboard_mapping__*` + `dashboard_check__*`; `nghiem_thu`: read +
`dashboard_timloi__tl_so_dashboard` + `dashboard_check__check_so_hien_thi`; mọi agent deny
`group:runtime`/`write`/`edit`/`apply_patch`, `fs.workspaceOnly=true`. (check/keo_file/mapping giữ như đợt 1.)
```bash
cd $AG
.venv/bin/python scripts/ap_tool_theo_agent.py          # CHẠY THỬ: diff đã che bí mật + dry-run schema OpenClaw
.venv/bin/python scripts/ap_tool_theo_agent.py --ap     # sao lưu openclaw.json.bak-tool-theo-agent-<giờ> -> ghi -> validate
```
Nếu bỏ qua `openclaw agents add` ở bước 6, script tự dựng mục agent mới theo khuôn mục `check` (khai ở
`_tao_neu_thieu` của file JSON) — chạy thử sẽ in "agent '<id>' CHƯA có -> dựng mục mới". Đường chính vẫn là bước 6.
Hoàn tác: lệnh `docker exec … cp -p <bản sao lưu> …/openclaw.json` script in ra sau khi ghi.

## 8. Restart gateway (nạp agent + MCP mới)
```bash
docker restart openclaw-openclaw-gateway-1
$OC openclaw agents list | grep -E 'quan_ly|tim_loi|nghiem_thu'
# thử gọi trực tiếp (không qua chat) — phải trả lời, không rơi về agent mặc định:
$OC openclaw agent --agent quan_ly --session-key agent:quan_ly:thu-cai-dat -m "phieu_bang_tong: còn phiếu nào?"
```

## 9. Tạo kênh + điền người duyệt
```bash
cd $API
.venv/bin/python tools/tao_kenh_agent_tong.py --env test           # chạy thử
.venv/bin/python tools/tao_kenh_agent_tong.py --env test --ap
.venv/bin/python tools/tao_kenh_agent_tong.py --env prod           # chạy thử
.venv/bin/python tools/tao_kenh_agent_tong.py --env prod --ap      # in channel_id + user_id (AI.Agent.03)
```
Điền đúng số script in ra (PROD):
- `$API/.env`: `PHIEU_NGUOI_DUYET=<user_id>` (TEST là 66; PROD tra theo email `AI.Agent.03@thinhcuong.com.vn`).
- `$AG/.env`: `PHIEU_CHANNEL_ID=<channel_id>`.
```bash
sudo systemctl restart tc-admin-api tc-bot-worker dashboard-mcp-phieu
```

## 10. Nghiệm thu (trong kênh "Xử lý lỗi dashboard")
1. Dán 2-3 dòng log Excel (vd `$AG/scripts/du_lieu_test/log_mau_xetai_so_gia.tsv`) → Agent tổng trả bảng phiếu
   (mã `PL-YYYYMMDD-NNN`, màn htxt0/htxt1, kỳ chuẩn hoá) và báo đã giao tim_loi; mỗi phiếu có một tin gốc trong kênh.
2. Trong ~10 phút: luồng mỗi phiếu có tin `[tim_loi] …` (loại, số đúng, số PROD, lệnh `DUYET PHAM VI …` gợi ý).
   Theo dõi: `ls -t $AG/memory/phieu/viec | head`; `tail $AG/memory/phieu/audit.jsonl`.
3. Người KHÁC `PHIEU_NGUOI_DUYET` gõ `HOLD PL-…` → API trả "Bạn không có quyền duyệt phiếu".
4. Người duyệt gõ `DUYET PHAM VI PL-… khoi=HTXT man=htxt1 ky=2026-07` → phiếu DA_DUYET → Agent tổng giao sua_loi →
   có file `$AG/memory/phieu/viec_sua/PL-….json` và tin "chờ IT nhận việc sửa (Claude Code bot-sua-loi)".
5. "cho cột IT fix lỗi" → Agent tổng đăng khối TSV 2 cột đúng thứ tự dòng đã dán.

## Hoàn tác nhanh
- Tắt Agent tổng ở chat: bỏ `ai_quan_ly_bot=quan_ly` khỏi `OPENCLAW_DOCKER_BOTS` + restart API/worker.
- Dừng việc nền: `sudo systemctl stop dashboard-mcp-phieu` (việc đang chạy → `gian_doan` khi khởi động lại; phiếu giữ
  trạng thái cũ).
- Gỡ MCP: `$OC openclaw mcp unset dashboard_phieu` (và `dashboard_timloi`).
- openclaw.json: chép lại bản `openclaw.json.bak-tool-theo-agent-<giờ>`.
- DB: `migrations/0093_bug_tickets.down.sql` (XOÁ bảng phiếu — sao lưu `COPY bug_tickets TO …` trước).

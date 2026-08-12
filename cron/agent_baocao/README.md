# Agent giám sát "tình trạng update file báo cáo ngày"

Bản SAO LƯU chỉ dẫn của agent openclaw `update_file_bao_cao_ngay`. Bản ĐANG CHẠY nằm ở

    /home/itadmin/.openclaw/agents/update_file_bao_cao_ngay/workspace/AGENTS.md

thư mục đó thuộc user `sysadmin` mode 700 nên sửa phải qua container:

```sh
docker exec -i openclaw-openclaw-gateway-1 sh -c \
  'cat > /home/node/.openclaw/agents/update_file_bao_cao_ngay/workspace/AGENTS.md' < AGENTS.md
```

Sửa xong thì đồng bộ lại vào đây (`docker exec ... cat > AGENTS.md`) — hai bản lệch nhau thì bản
trong container mới là bản thật.

## Nó đọc gì

Không parse log. Chỉ đọc 2 artifact JSON do cron sinh ra (xem `../cron_status.py`):

    ../logs/status_hqkdngay_daily_prod.json     (+ .jsonl lịch sử)
    ../logs/status_thuchi_daily_prod.json

## Chạy khi nào

Cron của openclaw, không phải crontab của máy:

```sh
docker exec openclaw-openclaw-gateway-1 openclaw cron list
docker exec openclaw-openclaw-gateway-1 openclaw cron runs --id 4d0d2b09-7910-4d75-b91e-43c7e94da3a8
```

Job `4d0d2b09-7910-4d75-b91e-43c7e94da3a8`: `15 17 * * *` giờ Asia/Ho_Chi_Minh (container chạy TZ=UTC
nên **bắt buộc** có `--tz`, thiếu là job nổ 00:15 giờ VN). Hiện `delivery: none` — chưa nối kênh gửi.
Nối kênh: `openclaw cron edit <id> --announce --channel <kênh>`.

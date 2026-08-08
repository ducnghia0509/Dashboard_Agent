#!/usr/bin/env bash
# Đồng bộ SKILL.md từ REPO (nguồn team sửa) -> WORKSPACE OpenClaw thực nạp.
# BẮT BUỘC chạy sau MỖI lần sửa agents/*/SKILL.md — OpenClaw KHÔNG tự đọc bản repo,
# nó nạp <state>/agents/<agent>/workspace/skills/*/SKILL.md (bản riêng, dễ đóng băng lệch).
# Phát hiện 2026-07-10: 3/4 SKILL workspace đóng băng từ 2-4/7 trong khi repo đã sửa nhiều lần.
#
# Sửa 2026-08-07 — bản trước KHÔNG chạy được trên máy này mà vẫn exit 0, nên "đã đồng bộ" là giả:
#   1. STATE_DIR mặc định '$HOME/openclaw/state/agents' KHÔNG tồn tại với user itadmin
#      (state thật: ~/.openclaw/agents). Vòng lặp in "bỏ qua" cả 4 agent rồi thoát êm.
#   2. Container tên 'openclaw-openclaw-gateway-1', không phải 'openclaw' -> restart luôn lỗi.
#   3. ~/.openclaw thuộc user sysadmin, itadmin KHÔNG ghi được từ host -> phải đi qua
#      'docker exec' (container chạy user node và mount đúng thư mục đó).
#   4. Thiếu workspace nay là LỖI (exit 2), không im lặng cho qua.
set -euo pipefail

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)/agents"
CONTAINER="${OPENCLAW_CONTAINER:-openclaw-openclaw-gateway-1}"
# Đường dẫn state NHÌN TỪ TRONG CONTAINER (host mount ~/.openclaw -> /home/node/.openclaw).
STATE_DIR="${OPENCLAW_STATE_DIR:-/home/node/.openclaw/agents}"
AGENTS="${OPENCLAW_AGENTS:-analyst orchestrator execute qa qa_chutich}"

if ! docker inspect "$CONTAINER" >/dev/null 2>&1; then
  echo "LỖI: không thấy container '$CONTAINER'. Đặt OPENCLAW_CONTAINER nếu tên khác." >&2
  exit 1
fi

changed=0
missing=0
for a in $AGENTS; do
  repo="$REPO_DIR/$a/SKILL.md"
  if [ ! -f "$repo" ]; then
    echo "$a: bỏ qua (repo không có SKILL.md)"
    continue
  fi
  ws=$(docker exec "$CONTAINER" sh -c \
        "ls $STATE_DIR/$a/workspace/skills/*/SKILL.md 2>/dev/null | head -1" || true)
  if [ -z "$ws" ]; then
    echo "$a: LỖI — workspace chưa có skill dir ($STATE_DIR/$a/workspace/skills/*/SKILL.md)" >&2
    missing=$((missing + 1))
    continue
  fi
  if docker exec "$CONTAINER" sh -c "cat '$ws'" 2>/dev/null | diff -q - "$repo" >/dev/null 2>&1; then
    echo "$a: đã khớp"
  else
    docker exec -i "$CONTAINER" sh -c "cat > '$ws'" < "$repo"
    echo "$a: ĐỒNG BỘ ($(wc -c <"$repo") byte -> $ws)"
    changed=1
  fi
done

if [ "$changed" = 1 ]; then
  echo "-> restart $CONTAINER để nạp skill mới..."
  docker restart "$CONTAINER" >/dev/null && echo "$CONTAINER restarted"
fi

# Thiếu workspace = agent đó KHÔNG nhận bản SKILL mới. Phải fail để người chạy biết, đừng lặp
# lại lỗi cũ (im lặng exit 0 trong khi thực chất chưa đồng bộ gì).
if [ "$missing" -gt 0 ]; then
  echo "LỖI: $missing agent chưa có workspace skill dir — CHƯA đồng bộ được." >&2
  exit 2
fi

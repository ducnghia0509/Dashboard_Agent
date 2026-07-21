#!/usr/bin/env bash
# Đồng bộ SKILL.md từ REPO (nguồn team sửa) -> WORKSPACE OpenClaw thực nạp.
# BẮT BUỘC chạy sau MỖI lần sửa agents/*/SKILL.md — OpenClaw KHÔNG tự đọc bản repo,
# nó nạp state/agents/<agent>/workspace/skills/*/SKILL.md (bản riêng, dễ đóng băng lệch).
# Phát hiện 2026-07-10: 3/4 SKILL workspace đóng băng từ 2-4/7 trong khi repo đã sửa nhiều lần.
set -euo pipefail
REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)/agents"
STATE_DIR="/home/sysadmin/openclaw/state/agents"
changed=0
for a in analyst orchestrator execute qa; do
  repo="$REPO_DIR/$a/SKILL.md"
  ws=$(ls "$STATE_DIR/$a/workspace/skills"/*/SKILL.md 2>/dev/null | head -1 || true)
  [ -f "$repo" ] || { echo "$a: bỏ qua (repo không có)"; continue; }
  [ -n "$ws" ]  || { echo "$a: bỏ qua (workspace chưa có skill dir)"; continue; }
  if diff -q "$ws" "$repo" >/dev/null 2>&1; then
    echo "$a: đã khớp"
  else
    cp "$repo" "$ws" && echo "$a: ĐỒNG BỘ ($(wc -c <"$ws")b)" && changed=1
  fi
done
if [ "$changed" = 1 ]; then
  echo "-> restart openclaw để nạp skill mới..."
  docker restart openclaw >/dev/null && echo "openclaw restarted"
fi

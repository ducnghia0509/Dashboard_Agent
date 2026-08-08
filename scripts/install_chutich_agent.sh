#!/usr/bin/env bash
# Tạo agent OpenClaw `qa_chutich`: workspace + entry trong openclaw.json.
#
# Idempotent, chạy lại được. KHÔNG restart gateway (bản đầu có restart và đó là sai — xem
# chú thích bước 3/3).
#
#   bash scripts/install_chutich_agent.sh            # xem trước, không ghi
#   bash scripts/install_chutich_agent.sh --apply    # thực sự tạo
#
# ~/.openclaw thuộc user sysadmin, itadmin không ghi được từ host -> mọi thao tác đi qua
# `docker exec` (container chạy user node và mount đúng thư mục đó).
set -euo pipefail

CONTAINER="${OPENCLAW_CONTAINER:-openclaw-openclaw-gateway-1}"
AGENT="qa_chutich"
SKILL_DIR_NAME="chutich"
STATE="/home/node/.openclaw"
REPO_SKILL="$(cd "$(dirname "$0")/.." && pwd)/agents/$AGENT/SKILL.md"
APPLY=0
[ "${1:-}" = "--apply" ] && APPLY=1

[ -f "$REPO_SKILL" ] || { echo "LỖI: không thấy $REPO_SKILL" >&2; exit 1; }
docker inspect "$CONTAINER" >/dev/null 2>&1 || {
  echo "LỖI: không thấy container '$CONTAINER'" >&2; exit 1; }

echo "Agent      : $AGENT"
echo "Container  : $CONTAINER"
echo "SKILL nguồn: $REPO_SKILL ($(wc -c <"$REPO_SKILL") byte)"
echo "Đã có sẵn  : $(docker exec "$CONTAINER" sh -c "[ -d $STATE/agents/$AGENT ] && echo CÓ || echo chưa")"
echo "Entry hiện có trong openclaw.json:"
docker exec "$CONTAINER" python3 -c "
import json
d=json.load(open('$STATE/openclaw.json'))
print('   ', sorted(d.get('agents',{}).get('entries',{})))
"

if [ "$APPLY" != 1 ]; then
  echo
  echo "XEM TRƯỚC — chưa ghi gì. Thêm --apply để thực hiện."
  echo "(Bản này KHÔNG restart gateway nữa — đăng ký bằng CLI openclaw, xem chú thích bước 3/3.)"
  exit 0
fi

echo
echo "1/3 Tạo workspace + nạp file nền RIÊNG..."
# KHÔNG copy AGENTS.md/SOUL.md từ agent 'qa' nữa. Bản đầu làm vậy và đã hỏng thật (08/08/2026):
# hai file đó là TEMPLATE MẶC ĐỊNH của OpenClaw, mô tả một trợ lý cá nhân có trí nhớ hằng ngày
# (`memory/YYYY-MM-DD.md`), lịch, email. Với câu hỏi mở như "hôm nay có gì bất thường?" — đúng
# câu Chủ tịch gõ đầu tiên mỗi sáng — agent bám khung đó thay vì bám SKILL, không thèm mở brief,
# rồi đi mời Chủ tịch cho xem lịch với email. 70/71 câu đạt nhưng hỏng đúng câu đầu tiên.
# Bản nền riêng nằm ở agents/qa_chutich/workspace/ trong repo, neo thẳng vào vai trò.
docker exec "$CONTAINER" sh -c "
  set -e
  mkdir -p '$STATE/agents/$AGENT/workspace/skills/$SKILL_DIR_NAME' \
           '$STATE/agents/$AGENT/agent' '$STATE/agents/$AGENT/sessions'
"
REPO_WS="$(cd "$(dirname "$0")/.." && pwd)/agents/$AGENT/workspace"
for f in AGENTS.md SOUL.md IDENTITY.md; do
  if [ -f "$REPO_WS/$f" ]; then
    docker exec -i "$CONTAINER" sh -c "cat > '$STATE/agents/$AGENT/workspace/$f'" < "$REPO_WS/$f"
    echo "   nạp $f ($(wc -c <"$REPO_WS/$f") byte)"
  fi
done
# TOOLS.md / USER.md: dùng bản trơn, không mang theo ví dụ camera/SSH của template gốc.
docker exec "$CONTAINER" sh -c "
  : > '$STATE/agents/$AGENT/workspace/TOOLS.md'
  : > '$STATE/agents/$AGENT/workspace/USER.md'
"

echo "2/3 Nạp SKILL.md..."
docker exec -i "$CONTAINER" sh -c \
  "cat > '$STATE/agents/$AGENT/workspace/skills/$SKILL_DIR_NAME/SKILL.md'" < "$REPO_SKILL"

echo "3/3 Đăng ký agent qua CLI openclaw..."
# KHÔNG sửa openclaw.json bằng tay. Đã thử và HỎNG (07/08/2026): gateway tự quản file này và
# lúc tắt sẽ ghi đè bằng bản trong bộ nhớ của nó — entry vừa thêm biến mất sạch sau restart,
# trong khi restart vẫn kịp làm gián đoạn cả prod lẫn coding. Dấu hiệu nhận biết cơ chế đó là
# loạt openclaw.json.bak / .bak.1..4 / .last-good do chính gateway sinh ra.
# CLI đi qua đúng đường quản lý config của gateway, ghi bền, và KHÔNG cần restart.
docker exec "$CONTAINER" openclaw agents add "$AGENT" \
  --workspace "$STATE/agents/$AGENT/workspace" \
  --agent-dir "$STATE/agents/$AGENT/agent" \
  --non-interactive --json

echo
echo "Kiểm tra entry đã ghi bền:"
docker exec "$CONTAINER" openclaw agents list --json \
  | python3 -c "import json,sys; print('   ', sorted(a['id'] for a in json.load(sys.stdin)))"

echo
echo "XONG. Kiểm tra tiếp:"
echo "  bash scripts/sync_skills.sh          # phải in '$AGENT: đã khớp'"
echo "  .venv/bin/python scripts/run_chutich_eval.py --limit 3"
echo
echo "Gỡ nếu cần (cũng không phải restart):"
echo "  docker exec $CONTAINER openclaw agents delete $AGENT"

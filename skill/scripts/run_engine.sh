#!/usr/bin/env bash
# 毕业论文格式化 Agent · 引擎启动器
# 自动定位 thesis_md2docx 引擎并加入 PYTHONPATH, 再调用 `python -m thesis_md2docx.main`
# 优先级: 环境变量 THESIS_MD2DOCX_REPO -> skill 父目录(即仓库根, 引擎在仓库根) -> skill 目录(打包时引擎若在 skill/ 下) -> WorkBuddy 插件路径
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"          # .../skill/scripts
SKILL_DIR="$(cd "$HERE/.." && pwd)"            # .../skill
REPO_ROOT="$(cd "$SKILL_DIR/.." && pwd)"       # 仓库根(thesis_md2docx 所在目录)
PY="${PYTHON:-python}"

CANDIDATES=(
  "${THESIS_MD2DOCX_REPO:-}"
  "$REPO_ROOT"
  "$SKILL_DIR"
  "$HOME/.workbuddy/plugins/marketplaces/my-experts/plugins/hit-thesis-formatter/skills/hitmd2docx"
)
REPO=""
for c in "${CANDIDATES[@]}"; do
  [ -z "$c" ] && continue
  if [ -d "$c/thesis_md2docx" ]; then REPO="$c"; break; fi
done
if [ -z "$REPO" ]; then
  echo "ERROR: 找不到 thesis_md2docx 引擎。请设置 THESIS_MD2DOCX_REPO 指向含 thesis_md2docx/ 的目录。" >&2
  exit 2
fi
export PYTHONPATH="$REPO:${PYTHONPATH}"
exec "$PY" -m thesis_md2docx.main "$@"

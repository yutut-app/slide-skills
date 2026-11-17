#!/bin/sh
# Mac / Linux：5本のスキルを ~/.claude/skills に**写して**入れる（リンクにしたいなら tools/link.sh）。
#   sh tools/install.sh
# 同じ名前のフォルダは**丸ごと入れ替える**（古いファイルが残ると、参照ファイルの中身が食い違う）。
set -eu
ROOT=$(cd "$(dirname "$0")/.." && pwd -P)
DEST="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
OLD="$HOME/.claude/skills-old"
python3 "$ROOT/tools/bundle_content.py" --check >/dev/null || { echo "**中身の規約の写しが古い。**先に python3 tools/bundle_content.py"; exit 1; }
mkdir -p "$DEST"
for d in "$ROOT"/skills/*/; do
  name=$(basename "$d")
  [ -L "${DEST}/${name}" ] && rm "${DEST}/${name}"
  rm -rf "${DEST:?}/${name}"
  cp -R "$d" "${DEST}/${name}"
  echo "入れた     ${name}"
done
for name in slider-craft slider-reskin create-slide; do
  if [ -e "${DEST}/${name}" ]; then mkdir -p "${OLD}"; mv "${DEST}/${name}" "${OLD}/"; echo "移した     ${name} → ${OLD}（使わない旧スキル。中身は残る）"; fi
done
python3 "$ROOT/tools/verify_install.py" "$DEST"

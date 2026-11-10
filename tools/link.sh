#!/bin/sh
# スキルを ~/.claude/skills に登録し直す。**git pull の後に毎回1回実行する。**
#
#   sh tools/link.sh          # 登録する（旧名のリンクを外す）
#   sh tools/link.sh --check  # 何もせず、状態だけ出す
#
# 外すのは「このリポジトリを指すリンク」だけ。**フォルダそのものは消さない**（警告だけ出す）。
set -eu
ROOT=$(cd "$(dirname "$0")/.." && pwd -P)
DEST="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
CHECK=0; [ "${1:-}" = "--check" ] && CHECK=1
LEGACY="slider-craft slider-reskin create-slide"
mkdir -p "$DEST"

for d in "$ROOT"/skills/*/; do
  name=$(basename "$d"); target="$ROOT/skills/${name}"; link="$DEST/${name}"
  if [ -L "${link}" ] && [ "$(readlink "${link}")" = "$target" ]; then
    echo "登録済み   ${name}"
  elif [ -e "${link}" ] && [ ! -L "${link}" ]; then
    echo "**要確認** ${name}: リンクでないフォルダがある（消さない。中身を見て移すか決める）"
  else
    [ $CHECK = 1 ] && { echo "未登録     ${name}"; continue; }
    ln -sfn "$target" "${link}"; echo "登録した   ${name}"
  fi
done

for name in $LEGACY; do
  link="$DEST/${name}"
  [ -e "${link}" ] || [ -L "${link}" ] || continue
  if [ -L "${link}" ]; then
    case "$(readlink "${link}")" in
      "$ROOT"/*|*/slide-skills/*)
        [ $CHECK = 1 ] && { echo "旧名が残る ${name}（--check なしで外す）"; continue; }
        rm "${link}"; echo "外した     ${name}（旧名。中身はリポジトリに残る）";;
      *) echo "**要確認** ${name}: 別の場所を指すリンク → $(readlink "${link}")";;
    esac
  else
    echo "**要確認** ${name}: 旧スキルのフォルダがある。使わないなら ~/.claude/skills の外へ移す（消さない）"
  fi
done

# 実体の無いリンク（名前を変えた後に残ったもの）
for link in "$DEST"/*; do
  [ -L "${link}" ] && [ ! -e "${link}" ] || continue
  case "$(readlink "${link}")" in "$ROOT"/*)
    [ $CHECK = 1 ] && { echo "切れたリンク $(basename "${link}")"; continue; }
    rm "${link}"; echo "外した     $(basename "${link}")（切れたリンク）";;
  esac
done

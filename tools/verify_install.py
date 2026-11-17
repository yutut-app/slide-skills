#!/usr/bin/env python3
"""入れた先のスキルが、正本と同じかを確かめる。install.sh / install.ps1 の最後に呼ぶ。

    python3 tools/verify_install.py [<入れた先>]     # 既定は ~/.claude/skills

**ファイルの一覧と中身（ハッシュ）を正本と比べる。**足りない・多い・中身が違う、を全部出す。
実績: フォルダを写して使った機械で、参照ファイル1本の中身が別のファイルになっていた。
"""
import hashlib
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEST = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else Path(os.environ.get("CLAUDE_SKILLS_DIR", Path.home() / ".claude" / "skills"))
SKIP = {"__pycache__", ".DS_Store"}


def files(base: Path):
    out = {}
    for p in base.rglob("*"):
        if p.is_file() and not (set(p.relative_to(base).parts) & SKIP) and not p.name.endswith(".pyc"):
            out[p.relative_to(base).as_posix()] = hashlib.md5(p.read_bytes()).hexdigest()
    return out


bad = 0
for src in sorted((ROOT / "skills").iterdir()):
    if not (src / "SKILL.md").is_file():
        continue
    dst = DEST / src.name
    if not (dst / "SKILL.md").is_file():
        print(f"**無い** {src.name}"); bad += 1; continue
    a, b = files(src), files(dst.resolve())
    missing = sorted(set(a) - set(b)); extra = sorted(set(b) - set(a))
    diff = sorted(k for k in set(a) & set(b) if a[k] != b[k])
    if missing or extra or diff:
        bad += 1
        print(f"**食い違い** {src.name}: 足りない {len(missing)}、多い {len(extra)}、中身が違う {len(diff)}")
        for k in (missing[:3] + extra[:3] + diff[:3]):
            print(f"  - {k}")
    else:
        print(f"一致       {src.name}（{len(a)} ファイル）")
for old in ("slider-craft", "slider-reskin", "create-slide"):
    if (DEST / old).exists():
        print(f"**旧スキルが残っている** {old}（使わない。登録先の外へ移す）"); bad += 1
print(f"検査: verify_install  対象 {DEST}  指摘 {bad}件  前提: 正本はこのリポジトリの skills/")
sys.exit(1 if bad else 0)

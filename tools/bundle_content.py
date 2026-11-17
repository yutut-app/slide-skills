#!/usr/bin/env python3
"""中身の規約（リポジトリ直下の正本）を、5スキルそれぞれの `references/content/` に写す。

    python3 tools/bundle_content.py          # 写す（commit と publish の前に必ず）
    python3 tools/bundle_content.py --check  # 写しが正本と同じかだけ見る（違えば終了1）

**なぜ写すか：**`skills/` の中のフォルダだけをコピーして使う機械がある。正本をリポジトリ直下にしか
置かないと、その機械では規約が1つも届かない（実績: 判定表に「無し」と出たまま作業が進んだ）。
**正本はリポジトリ直下 `assets/content/` と `assets/project-layout.md`。写しを直接直さない。**
"""
import hashlib
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = [*sorted((ROOT / "assets" / "content").glob("*.md")), ROOT / "assets" / "project-layout.md"]
check = "--check" in sys.argv
bad = 0
for skill in sorted((ROOT / "skills").iterdir()):
    if not (skill / "SKILL.md").is_file():
        continue
    dst_dir = skill / "references" / "content"
    dst_dir.mkdir(parents=True, exist_ok=True)
    names = {s.name for s in SRC}
    for old in dst_dir.glob("*.md"):            # 正本から消えた写しを残さない
        if old.name not in names:
            bad += 1 if check else 0
            if not check:
                old.unlink()
    for s in SRC:
        d = dst_dir / s.name
        same = d.exists() and hashlib.md5(d.read_bytes()).digest() == hashlib.md5(s.read_bytes()).digest()
        if same:
            continue
        if check:
            bad += 1
            print(f"**食い違い** {d.relative_to(ROOT)}")
        else:
            shutil.copy2(s, d)
print(f"検査: bundle_content  対象 {len(SRC)}件×スキル  指摘 {bad}件  前提: 正本は assets/content/ と assets/project-layout.md")
sys.exit(1 if bad else 0)

#!/usr/bin/env python3
"""スキル間で同じ物を置く道具が、**1字違わず同じか**を確かめる。publish の前に必ず通す。

    python3 tools/sync_check.py          # 食い違いがあれば終了1
    python3 tools/sync_check.py --fix <正本のスキル名>   # その版で他を上書きする

**同じ名前の .py が2つ以上のスキルにあれば、同じ物であるべき**とみなす。
食い違いを残すと、どのスキルから呼んだかで結果が変わる（実績: checked.py の2版が混在していた）。
"""
import hashlib
import shutil
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
groups = defaultdict(list)
for p in sorted(ROOT.glob("skills/*/scripts/*.py")):
    groups[p.name].append(p)

bad = 0
fix = sys.argv[sys.argv.index("--fix") + 1] if "--fix" in sys.argv else None
for name, paths in sorted(groups.items()):
    if len(paths) < 2:
        continue
    digests = {p: hashlib.md5(p.read_bytes()).hexdigest() for p in paths}
    if len(set(digests.values())) == 1:
        continue
    src = next((p for p in paths if p.parts[-3] == fix), None) if fix else None
    if src:
        for p in paths:
            if p != src:
                shutil.copy2(src, p)
        print(f"揃えた   {name}（正本: {fix}）")
        continue
    bad += 1
    print(f"**食い違い** {name}: " + ", ".join(f"{p.parts[-3]}={d[:7]}" for p, d in digests.items()))

n = sum(1 for v in groups.values() if len(v) > 1)
print(f"検査: sync_check  対象 {n}種  指摘 {bad}件  前提: 2スキル以上に同じ名前で置かれた .py")
sys.exit(1 if bad else 0)

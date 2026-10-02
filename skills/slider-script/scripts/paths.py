#!/usr/bin/env python3
"""案件ルートと、出力先の既定を決める。**各スキルの scripts に同じ物を置く**
（`checked.py` と同じ扱い。スキルを1つだけ配った環境でも動かすため）。

正本は `assets/project-layout.md`。**ここに決めごとを書き足さない。**

    from paths import project_root, out_dir
    root = project_root(args.project, hint=Path(args.html))
    outdir = args.outdir or out_dir(root, "qa")

**最初に絶対パスへ解決する。**相対パスのまま扱うと、呼ぶ場所で結果が変わる
（別の道具で、在るものを「無い」と報告する事故が実際に起きた）。
"""

from pathlib import Path

MARK = "project.md"      # 案件ルートの目印
LATEST = "latest"        # いま作業している版。**見るのは常にここ**


def project_root(explicit=None, hint=None) -> Path:
    """`--project` > `project.md` を上に辿る > 今いる場所。**必ず絶対パス。**"""
    if explicit:
        return Path(explicit).expanduser().resolve()
    start = Path(hint).expanduser().resolve() if hint else Path.cwd().resolve()
    if start.is_file():
        start = start.parent
    for d in [start, *start.parents]:
        if (d / MARK).is_file():
            return d
    return Path.cwd().resolve()


def out_dir(root: Path, name: str, version: str = LATEST) -> Path:
    """`<案件>/out/<版>/<用途>`。**用途ごとに1つ**（別名を増やさない）。"""
    d = root / "out" / version / name
    d.mkdir(parents=True, exist_ok=True)
    return d


def where(root: Path) -> str:
    """報告に貼る1行。**どこに出したかを毎回書く。**"""
    return f"出力: {root / 'out' / LATEST}（案件: {root.name}）"

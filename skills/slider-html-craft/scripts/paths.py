#!/usr/bin/env python3
"""案件ルートと、出力先の既定を決める。**各スキルの scripts に同じ物を置く**
（`checked.py` と同じ扱い。スキルを1つだけ配った環境でも動かすため）。

正本は `assets/project-layout.md`。**ここに決めごとを書き足さない。**

    from paths import project_root, out_dir
    python3 scripts/paths.py --layout     # 正本（project-layout.md）の実在するパスを出す
    python3 scripts/paths.py --content    # 中身の規約（assets/content/）の実在するパスを出す
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


def layout_doc():
    """正本 `assets/project-layout.md` の実在するパス。無ければ None。

    **スキルのフォルダではなく、リポジトリ直下にある。**登録はスキルのフォルダへのリンクなので、
    `assets/project-layout.md` をスキルから見た相対で探すと見つからない
    （実績: 試験の作業エージェントが2本とも「スキル一式に無い」と報告した）。
    リンクを辿った実体（resolve）から上がって探す。
    """
    here = Path(__file__).resolve()
    for up in here.parents:
        cand = up / "assets" / "project-layout.md"
        if cand.is_file():
            return cand
    return None


def content_dir():
    """中身の規約（意図・骨子・図・日本語・数値・用語・先方の目）の実在する場所。無ければ None。

    **2つの craft（html / pptx）が同じ規約を読む。**写しを持つと食い違うので、
    リポジトリ直下の `assets/content/` に1つだけ置く。探し方は `layout_doc()` と同じ。
    """
    here = Path(__file__).resolve()
    for up in here.parents:
        cand = up / "assets" / "content"
        if (cand / "40_japanese.md").is_file():
            return cand
    return None


if __name__ == "__main__":
    import sys
    if "--content" in sys.argv:
        d = content_dir()
        print(d if d else "**無し**（スキルのフォルダだけが置かれた環境）。中身の規約が読めない。リポジトリを clone して `sh tools/link.sh` で登録する")
        sys.exit(0 if d else 1)
    if "--layout" in sys.argv:
        d = layout_doc()
        print(d if d else "**無し**（スキルのフォルダだけが置かれた環境）。要点だけ守る: 受け取りは `<案件>/in/`、出力は `<案件>/out/latest/`、名前は固定、日時を名前に入れない")
        sys.exit(0 if d else 1)
    print(where(project_root()))

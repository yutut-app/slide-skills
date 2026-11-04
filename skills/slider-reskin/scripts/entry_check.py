#!/usr/bin/env python3
"""入口の判定。**作業の最初に必ず通し、出てきた表を最初の出力にする。**

    python3 scripts/entry_check.py [<受け取ったファイル> ...] [--project <案件>]

前提（HTML の正本、テンプレート、道具、隣のスキル、版の扱い）を機械で確かめ、
**どの入口で進むか**と、**その入口で読む参照ファイル**を出す。
同じ物を slider-deck と slider-reskin の scripts に置く（`paths.py` と同じ扱い）。

**この表を出さずに作業を始めない。**実績として、前提を確かめずに進み、
使えないと分かったのが利用者に聞かれた後だった（HTML も書式の登録も無い案件）。

| 判定 | 入口 |
|---|---|
| スライドの HTML がある | (1) HTML から変換する（slider-deck） |
| HTML が無く、pptx がある | (2) 既存の pptx を土台に、必要な箇所だけ直す（slider-deck の差分修正） |
| どちらも無い | **作らない。**何を受け取るべきかを伝える |

書式だけを既存の pptx に合わせたいときは slider-reskin。**意図は機械では分からないので、表で聞く。**
"""

import argparse
import os
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
from paths import project_root

HERE = Path(__file__).resolve()
REPO = HERE.parents[3]                       # <repo>/skills/<skill>/scripts/entry_check.py
SKILLS_HOME = Path.home() / ".claude" / "skills"
_REV = re.compile(r"(?i)(?:^|[_\-\s])rev\.?\s*\d+")

REFS = {
    "1": ["skills/slider-deck/references/05_receive.md",
          "skills/slider-deck/references/55_revise.md",
          "skills/slider-deck/references/60_qa.md"],
    "2": ["skills/slider-deck/references/00_entry.md",
          "skills/slider-deck/references/55_revise.md",
          "skills/slider-deck/references/60_qa.md",
          "skills/slider-deck/references/07_official.md"],
}

# 入口(2) で「書式だけ合わせたい」ときに読む（slider-reskin の入口(2)）
RESKIN_REFS = ["skills/slider-reskin/references/30_from_pptx.md",
               "skills/slider-deck/references/07_official.md"]


def is_slide_html(p: Path) -> bool:
    if p.name == "print.html":
        return True
    try:
        head = p.read_text(encoding="utf-8", errors="ignore")[:200000]
    except OSError:
        return False
    return 'class="slide' in head or "class='slide" in head


_META_TPL = re.compile(r"""<meta[^>]+name=["']template["'][^>]+content=["']([^"']+)["']""", re.I)


def meta_templates(html):
    """HTML の <meta name="template"> が指す pptx。**登録していなくても、変換で使われる。**
    実績: in/template.pptx を meta で指していたのに、判定表は「テンプレート無し」と出した。"""
    out = []
    for h in html:
        try:
            head = h.read_text(encoding="utf-8", errors="ignore")[:20000]
        except OSError:
            continue
        for m in _META_TPL.findall(head):
            t = (h.parent / m).resolve()
            if t.exists() and t not in out:
                out.append(t)
    return out


def templates(root: Path):
    """登録済みのテンプレート。環境変数 → 案件の上の階層 → このリポジトリ。"""
    roots = []
    if os.environ.get("SLIDE_TEMPLATE_DIR"):
        roots.append(Path(os.environ["SLIDE_TEMPLATE_DIR"]))
    for up in [root, *root.parents][:5]:
        roots.append(up / "assets" / "templates" / "deck")
    roots.append(REPO / "assets" / "templates" / "deck")
    seen = []
    for r in roots:
        if r.is_dir():
            for m in sorted(r.glob("*/template-manifest.md")):
                if m.parent.name not in seen:
                    seen.append(m.parent.name)
    return seen


def tools():
    """画像化・変換に使う道具。**無いものは、何で代わりにするかと一緒に出す。**"""
    out = []
    soffice = shutil.which("soffice") or shutil.which("libreoffice") or next(
        (p for p in ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
         if Path(p).exists()), None)
    out.append(("LibreOffice（pptx を画像に）", "有り" if soffice else "**無し**"))
    out.append(("pdftoppm（PDF を画像に）", "有り" if shutil.which("pdftoppm") else "**無し**"))
    browser = any(Path(p).exists() for p in [
        "/Applications/Google Chrome.app", "/Applications/Microsoft Edge.app"]) or \
        shutil.which("chrome") or shutil.which("msedge") or shutil.which("google-chrome")
    out.append(("Chrome / Edge（HTML を画像に）", "有り" if browser else "**無し**"))
    if sys.platform.startswith("win"):
        try:
            import win32com.client  # noqa: F401
            com = "有り（PowerPoint 本体の起動は**未確認**）"
        except ImportError:
            com = "**無し**（pywin32 が無い）"
        out.append(("PowerPoint COM（Windows）", com))
    return out, bool(soffice)


def neighbor_skills():
    names = ["slider-craft", "slider-deck", "slider-reskin", "slider-script"]
    return [(n, "有り" if (SKILLS_HOME / n / "SKILL.md").exists() else "**無し**")
            for n in names]


def main():
    ap = argparse.ArgumentParser(description="入口の判定。作業の最初に必ず通す")
    ap.add_argument("files", nargs="*", help="受け取ったファイル（HTML / pptx）")
    ap.add_argument("--project", help="案件のディレクトリ。省くと project.md を上に辿る")
    args = ap.parse_args()

    hint = args.files[0] if args.files else None
    root = project_root(args.project, hint=hint)

    cand = [Path(f).expanduser().resolve() for f in args.files]
    for sub in ["in", "out/latest", "."]:
        d = root / sub
        if d.is_dir():
            cand += [p for p in d.glob("*") if p.is_file()]
    cand = sorted(set(cand))
    html = [p for p in cand if p.suffix.lower() == ".html" and is_slide_html(p)]
    pptx = [p for p in cand if p.suffix.lower() in (".pptx", ".potx")
            and not p.name.startswith("~$")]
    revs = [p.name for p in cand if _REV.search(p.stem)]
    tpls = templates(root)
    meta_tpls = meta_templates(html)
    tpls += [f"{t.name}（HTML の meta が指す）" for t in meta_tpls]
    # HTML と別に、テンプレートとして指されていない pptx がある＝書式の元かもしれない
    other_pptx = [p for p in pptx if p not in meta_tpls]
    tool_rows, can_render = tools()
    skills = neighbor_skills()
    craft = dict(skills).get("slider-craft") == "有り"

    if html:
        entry, why = "1", "スライドの HTML がある"
    elif pptx:
        entry, why = "2", "HTML が無く、pptx がある"
    else:
        entry, why = "—", "HTML も pptx も見つからない"

    print("## 入口の判定\n")
    print("| 確かめたこと | 結果 |")
    print("|---|---|")
    print(f"| スライドの HTML | {len(html)} 件{'' if html else '（**無し**）'} |")
    print(f"| pptx | {len(pptx)} 件{'' if pptx else '（無し）'} |")
    print(f"| テンプレート（登録済み・HTML の指定） | {', '.join(tpls) if tpls else '**無し**'} |")
    for name, st in tool_rows:
        print(f"| {name} | {st} |")
    for name, st in skills:
        print(f"| スキル {name} | {st} |")
    print(f"| 版の名前（Rev 等）の形跡 | {', '.join(revs[:5]) if revs else '無し'} |")

    print(f"\n**入口：{entry}**（{why}）\n")
    if entry == "1":
        print("- HTML から pptx に変換し、ずれは HTML に差し戻す")
        if other_pptx:
            print(f"- **HTML と別に pptx がある**（{', '.join(p.name for p in other_pptx[:3])}）。"
                  "**その書式（マスター・レイアウト）に載せたいなら slider-reskin の入口(2)**。そのとき読むのは "
                  + ", ".join(f"`{r}`" for r in RESKIN_REFS) + "（意図を下で聞く）")
    elif entry == "2":
        print("- **既存の pptx が正本。最初から差分修正モード**（作り直さない。必要な箇所だけ直す）")
        print("- 上書き・削除の前に、利用者の手直しを差分で確かめて退避する")
        print("- **書式だけを合わせたいなら slider-reskin**（意図を下で聞く）。"
              "そのとき読むのは " + ", ".join(f"`{r}`" + ("" if (REPO / r).exists() else "（**まだ無い**）")
                                           for r in RESKIN_REFS))
    else:
        print("- **作らない。**HTML 一式か、土台にする pptx を受け取る")

    if entry in REFS:
        print("\n**この入口で読む参照ファイル**（読んでから手を動かす）\n")
        for r in REFS[entry]:
            mark = "" if (REPO / r).exists() else "（**まだ無い**）"
            print(f"- `{r}`{mark}")

    print("\n**このスキルが扱わないもの**：構成・文言・図解の選定（slider-craft の担当）。"
          + ("" if craft else "**この機械には slider-craft が無い。**構成を良くしたい依頼には応えられない"))

    print("\n**まとめて確認すること**（ここで1回だけ止まる）\n")
    print("1. 入口はこれでよいか（書式だけ合わせたい／中身から作りたい、なら違う入口）"
          + ("。**HTML と別の pptx は、書式の元か、参考か**" if entry == "1" and other_pptx else ""))
    print("2. **版の扱い**：別名で出すか、同じ版に上書きするか"
          + ("（Rev の形跡あり。案件の運用を確かめる）" if revs else ""))
    if not craft:
        print("3. 構成・文言の直しが要るときの扱い（この機械では扱えない）")
    if not can_render:
        print(f"{4 if not craft else 3}. 画像化の道具が無い。**何で代わりに見るか**"
              "（PowerPoint で書き出す等）。代わりが無ければ、見た目は**検査していない**と明記する")

    findings = 0 if entry != "—" else 1
    return checked.summary("entry_check", len(html) + len(pptx), "入力", findings, {
        "案件": root.name,
        "入口": entry,
        "テンプレート": f"{len(tpls)} 件",
        "画像化": "可" if can_render else "**不可（代替が要る）**",
    })


if __name__ == "__main__":
    sys.exit(main())

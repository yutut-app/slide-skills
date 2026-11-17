#!/usr/bin/env python3
"""入口の判定。**5つのスライドスキルのどれでも、作業の最初に必ず通し、出てきた表を最初の出力にする。**

    python3 scripts/entry_check.py [<受け取ったファイル> ...] [--project <案件>] [--want html|pptx|convert|check|script]

**同じ物を5スキルの scripts に置く**（`paths.py` と同じ扱い。一致は `tools/sync_check.py` で見る）。
どのスキルが先に呼ばれても**同じ表**が出て、担当が違えば正しいスキルへ渡す。

`--want` は依頼の言葉から決める（下の表）。**言葉で決まらないときは付けない。**表が聞く。

| 依頼の言葉 | --want | 担当 |
|---|---|---|
| 「HTMLで作って」「構成を考えて」「このHTMLを直して」 | html | slider-html-craft |
| 「HTMLをパワポにして」「pptxに変換」「テンプレート／客先書式に載せて」 | convert | slider-html-reskin |
| 「パワポで作って」「前回の会議資料をベースに」「このパワポを直して」 | pptx | slider-pptx-craft |
| 「検査して」「崩れていないか」「渡す前に確認」「差分を出して」 | check | slider-deck |
| 「カンペ」「台本」「ノートに入れて」 | script | slider-script |

**この表を出さずに作業を始めない。**実績として、呼び名（「自作パワポスキル」）から担当を推測して外し、
使っていない手順を「使った」と記録した。別の実績では、古い版のスキルが動いていたのに誰も気づかなかった。
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
from paths import content_dir, layout_doc, project_root

HERE = Path(__file__).resolve()
SELF = HERE.parents[1].name                  # <repo>/skills/<skill>/scripts/entry_check.py
REPO = HERE.parents[3]
SKILLS_HOME = Path(os.environ.get("CLAUDE_SKILLS_DIR", Path.home() / ".claude" / "skills"))
_REV = re.compile(r"(?i)(?:^|[_\-\s])rev\.?\s*\d+")

SKILLS = ["slider-html-craft", "slider-html-reskin", "slider-pptx-craft", "slider-deck", "slider-script"]
LEGACY = {"slider-craft": "slider-html-craft", "slider-reskin": "slider-html-reskin",
          "create-slide": "slider-html-craft / slider-pptx-craft"}
WANT = {"html": "slider-html-craft", "convert": "slider-html-reskin", "pptx": "slider-pptx-craft",
        "check": "slider-deck", "script": "slider-script"}

# 担当ごとに最初に読むもの。**表に出たものを読んでから手を動かす。**
REFS = {   # 担当のスキルのフォルダから見たパス
    "slider-html-craft": ["references/content/00_route.md", "references/content/10_intent.md",
                          "references/content/20_story.md", "references/48_html.md"],
    "slider-html-reskin": ["references/05_receive.md", "references/30_from_pptx.md",
                           "references/content/07_official.md"],
    "slider-pptx-craft": ["references/55_revise.md", "references/content/07_official.md",
                          "references/content/10_intent.md", "references/content/20_story.md",
                          "references/content/40_japanese.md"],
    "slider-deck": ["references/60_qa.md", "references/content/80_audience.md"],
    "slider-script": ["references/10_script.md", "references/content/10_intent.md"],
}

_MEETING = re.compile(r"(?i)(第\s*\d+\s*回|mtg\s*\d+|meeting|会議)")


def skill_dir(name):
    """隣のスキルのフォルダ。**同じ親の下 → 登録先**の順に探す（フォルダだけ写した機械でも引ける）。"""
    for d in (HERE.parents[2] / name, SKILLS_HOME / name):
        if (d / "SKILL.md").is_file():
            return d
    return None


def ref_problem(path: Path):
    """参照ファイルが無い・中身が別のファイル、なら理由。見出し（1行目）が `# <名前>` で始まるかで見る。
    実績: 写したときに取り違え、`55_revise.md` の中身が別のファイルだった。"""
    if not path.is_file():
        return "無い"
    if path.name == "project-layout.md":
        return None
    try:
        head = path.read_text(encoding="utf-8", errors="ignore").splitlines()[0]
    except (OSError, IndexError):
        return "読めない"
    return None if head.startswith(f"# {path.stem}") else f"中身が別のファイル（1行目: {head[:30]}）"


def lock_files(dirs):
    """Office のロックファイル（~$）と、元のファイルに書き込めるか。**ロックファイルは目安。**
    実績: 古いロックファイルが残っていて、開いているかを判断できなかった。"""
    rows = []
    for d in dirs:
        for lk in sorted(d.glob("~$*")):
            stem = lk.name[2:]
            real = next((p for p in d.iterdir() if p.name.endswith(stem) and not p.name.startswith("~$")), None)
            if real is None:
                rows.append((lk.name, "元のファイルが無い（古い残り）"))
                continue
            try:
                with open(real, "r+b"):
                    pass
                rows.append((lk.name, f"{real.name} に書き込める（**開いていない可能性が高い**。開いていないか利用者に確かめる）"))
            except OSError:
                rows.append((lk.name, f"**{real.name} に書き込めない（開いている）**。閉じてもらう"))
    return rows


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
    """HTML の <meta name="template"> が指す pptx。**登録していなくても、変換で使われる。**"""
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


def skill_version(skill_dir: Path) -> str:
    try:
        head = (skill_dir / "SKILL.md").read_text(encoding="utf-8")[:4000]
    except OSError:
        return "—"
    m = re.search(r'version:\s*"?([0-9.]+)"?', head)
    return m.group(1) if m else "—"


def repo_date() -> str:
    """このスキルの版の日付。**古い版が動いていないかを、表を見た人が気づけるようにする。**"""
    try:
        r = subprocess.run(["git", "-C", str(REPO), "log", "-1", "--format=%cs %h"],
                           capture_output=True, text=True, timeout=5)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        pass
    try:
        return datetime.fromtimestamp((HERE.parents[1] / "SKILL.md").stat().st_mtime).strftime("%Y-%m-%d") + "（ファイルの日付）"
    except OSError:
        return "不明"


def registered():
    rows = []
    for n in SKILLS:
        d = skill_dir(n)
        if d:
            rows.append((n, f"有り（版 {skill_version(d.resolve())}）"))
        else:
            rows.append((n, "**無い**（README「登録と更新」で5本とも入れる）"))
    legacy = [(n, LEGACY[n]) for n in LEGACY if (SKILLS_HOME / n).exists() or (SKILLS_HOME / n).is_symlink()]
    return rows, legacy


def decide(want, html, pptx):
    """担当を決める。**言葉で決まらず、材料でも決まらないときは None（聞く）。**"""
    if want:
        return WANT[want], f"依頼の言葉（--want {want}）"
    if html and not pptx:
        return None, "HTML がある。**中身を直すのか（html-craft）、pptx にするのか（html-reskin）を聞く**"
    if pptx and not html:
        return "slider-pptx-craft", "HTML が無く pptx がある。前の資料を土台に作る・直す。**検査だけなら slider-deck**"
    if html and pptx:
        return None, "HTML と pptx の両方がある。**どちらが正本かを聞く**"
    return None, "手元に何も無い。**HTML で作るか、pptx で作るかを聞く**（外すと全工程がやり直し）"


def main():
    ap = argparse.ArgumentParser(description="入口の判定。5スキル共通。作業の最初に必ず通す")
    ap.add_argument("files", nargs="*", help="受け取ったファイル（HTML / pptx）")
    ap.add_argument("--project", help="案件のディレクトリ。省くと project.md を上に辿る")
    ap.add_argument("--want", choices=sorted(WANT), help="依頼の言葉から決まる担当（決まらなければ付けない）")
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
    pptx = sorted([p for p in cand if p.suffix.lower() in (".pptx", ".potx")
                   and not p.name.startswith("~$")], key=lambda p: p.stat().st_mtime, reverse=True)
    revs = [p.name for p in cand if _REV.search(p.stem)]
    tpls = templates(root)
    meta_tpls = meta_templates(html)
    tpls += [f"{t.name}（HTML の meta が指す）" for t in meta_tpls]
    tool_rows, can_render = tools()
    reg, legacy = registered()
    owner, why = decide(args.want, html, pptx)

    # **入れ方の不備は、ここで止める。**無しを表の1行で出すだけでは、作業がそのまま進んだ（実績）
    stops = []
    if not (HERE.parents[1] / "references" / "content" / "40_japanese.md").is_file():
        stops.append(f"このスキル（{SELF}）に中身の規約（references/content/）が入っていない")
    if owner:
        od = skill_dir(owner)
        if od is None:
            stops.append(f"担当のスキル {owner} が入っていない")
        else:
            for r in REFS.get(owner, []):
                why_bad = ref_problem(od / r)
                if why_bad:
                    stops.append(f"{owner}/{r} が{why_bad}")
    locks = lock_files([root / "in", root / "out" / "latest", root])
    meet = sorted({p.name for p in root.iterdir() if p.is_dir() and _MEETING.search(p.name)}) if root.is_dir() else []

    print("## 入口の判定\n")
    print("| 確かめたこと | 結果 |")
    print("|---|---|")
    print(f"| **いま動いているスキル** | {SELF}（版 {skill_version(HERE.parents[1])}、{repo_date()}） |")
    for name, st in reg:
        print(f"| スキル {name} | {st} |")
    for name, new in legacy:
        print(f"| **旧スキル {name}** | **登録が残っている。使わない**（後継: {new}）。`sh tools/link.sh` で外す（中身は消えない） |")
    print(f"| スライドの HTML | {len(html)} 件{'' if html else '（無し）'} |")
    if pptx:
        newest = pptx[0]
        print(f"| pptx（新しい順） | {', '.join(p.name for p in pptx[:5])}。**最新は {newest.name}**"
              f"（{datetime.fromtimestamp(newest.stat().st_mtime):%m-%d %H:%M}） |")
    else:
        print("| pptx | 0 件（無し） |")
    print(f"| テンプレート（登録済み・HTML の指定） | {', '.join(tpls) if tpls else '無し'} |")
    for name, st in tool_rows:
        print(f"| {name} | {st} |")
    lay, con = layout_doc(), content_dir()
    print(f"| 置き場所の決めごと | {lay if lay else '**無し**（要点: 受け取りは in/、出力は out/latest/、名前は固定）'} |")
    print(f"| 中身の規約（assets/content） | {con if con else '**無し**。リポジトリを clone して登録する'} |")
    print(f"| 版の名前（Rev 等）の形跡 | {', '.join(revs[:5]) if revs else '無し'} |")
    print(f"| 会議ごとのフォルダ | {', '.join(meet[:5]) if meet else '見つからない'} |")
    for name, st in locks:
        print(f"| 開いているファイル（{name}） | {st} |")

    if stops:
        print("\n## **止まる：スキルの入れ方に不備がある**\n")
        for x in stops:
            print(f"- {x}")
        print("\n**このまま作業を始めない。**利用者に上の不備を伝え、スキルを入れ直してもらう"
              "（README「登録と更新」：`skills` の中のフォルダを5本とも写し直す）。")
        checked.summary("entry_check", len(html) + len(pptx), "入力", len(stops), {"止まる": f"{len(stops)} 件"})
        return 3

    print(f"\n**担当：{owner or '未決（下で聞く）'}**（{why}）\n")
    if owner and owner != SELF:
        print(f"- **このスキル（{SELF}）の担当ではない。{owner} に渡す。**この表を貼って、そのスキルの手順0から始める")
    if owner == "slider-pptx-craft" and pptx:
        print(f"- **最初に、人が手直しした最新版（{pptx[0].name}）と自分が前に出したものの差を一覧にする**（`55_revise.md` 0章）")
        print("- **直す前に、修正案の表を出して止まる。**承認の前に書き込まない")

    refs = REFS.get(owner) if owner else None
    if refs:
        print("\n**最初に読む参照ファイル**（読んでから手を動かす）\n")
        for r in refs:
            print(f"- `{owner}/{r}`")

    print("\n**まとめて確認すること**（ここで1回だけ止まる。担当のスキルの次の止まる所と、材料がそろえば1回にまとめてよい）\n")
    n = 1
    if not owner:
        print(f"{n}. **何を作るか**：HTML で作る（slider-html-craft）／pptx で作る（slider-pptx-craft）"
              + "／HTML を pptx にする（slider-html-reskin）／検査だけ（slider-deck）"); n += 1
    if owner != "slider-deck":   # 検査は書き込まないので、版の扱いは聞かない
        print(f"{n}. **版の扱い**：別に出すか、同じ版に上書きするか" + ("（Rev の形跡あり。案件の運用に従う）" if revs else "")); n += 1
    if not can_render:
        print(f"{n}. 画像化の道具が無い。**何で代わりに見るか**。代わりが無ければ、見た目は**検査していない**と明記する"); n += 1
    if meet:
        print(f"{n}. **会議で使うもの（pptx・カンペ・画像）の置き場**：会議ごとのフォルダ（{meet[0]} など）に最初から置くか"); n += 1
    if legacy:
        print(f"{n}. 旧スキルの登録が残っている。外してよいか（`sh tools/link.sh`。中身は消えない）"); n += 1

    print("\n**完了報告には「使ったスキルと手順」を、各道具のサマリ行（`検査: ...`）付きで書く。**"
          "呼んだが使わなかった手順は「使わなかった」と書く。")

    findings = 0 if owner else 1
    return checked.summary("entry_check", len(html) + len(pptx), "入力", findings, {
        "案件": root.name,
        "担当": owner or "未決",
        "旧スキル": f"{len(legacy)} 件",
        "画像化": "可" if can_render else "**不可（代替が要る）**",
    })


if __name__ == "__main__":
    sys.exit(main())

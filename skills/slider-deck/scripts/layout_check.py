#!/usr/bin/env python3
"""案件の配置を機械で検査する。**成果物を人に渡す前に通す。**

    python3 skills/slider-deck/scripts/layout_check.py <案件>

正本は `assets/project-layout.md`。**決めごとはここに書かない。**

| 見るもの | 落ちる条件 |
|---|---|
| 案件の目印 | `project.md` が無い（**案件ルートが決まらない**） |
| 直下のフォルダ | 決めた以外のフォルダがある（散らばりの始まり） |
| `out/latest/` | 無い。**見る場所が決まらない** |
| 同じ用途の別名 | `qa` と `qa_html` のように、同じ用途が2つ以上ある |
| 日時入りの名前 | `…_20260914-1649.pptx` のような名前がある（後から探せない） |
| 版のフォルダ | `out/` の下に `vNN` 以外（`latest` を除く）がある |

**指摘があれば異常終了する**（終了コード 1）。
"""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
from paths import LATEST, MARK

TOP = {"in", "out", "work"}                 # 直下に置いてよいフォルダ
USES = ["qa", "charts"]                     # out/<版>/ の下の用途
_DATE = re.compile(r"(?:^|[_-])20\d{6}(?:[-_]?\d{3,6})?(?=[._-]|$)")
_SIBLING = re.compile(r"^(qa|charts|deck|stage\d*)[_a-z0-9]*$")


def main():
    ap = argparse.ArgumentParser(description="案件の配置を機械で検査する")
    ap.add_argument("project", nargs="?", default=".", help="案件のディレクトリ")
    args = ap.parse_args()
    root = Path(args.project).expanduser().resolve()   # **最初に絶対パスへ**

    findings, notes = [], []
    if not (root / MARK).is_file():
        findings.append(("—", "案件の目印が無い",
                         f"**`{MARK}` を置く。**無いと道具が案件ルートを決められない"))

    dirs = sorted(d.name for d in root.iterdir() if d.is_dir() and not d.name.startswith("."))
    for d in dirs:
        if d not in TOP:
            findings.append((d, "直下に決めていないフォルダ",
                             f"**`in` / `out` / `work` のどれかに入れる。**"
                             f"（{', '.join(sorted(TOP))} 以外は置かない）"))

    # 同じ用途の別名（qa と qa_html が並ぶ、など）
    fam = {}
    for d in dirs:
        m = _SIBLING.match(d)
        if m:
            fam.setdefault(m.group(1), []).append(d)
    for base, names in fam.items():
        if len(names) > 1:
            findings.append(("/".join(names), "同じ用途が別名で並ぶ",
                             f"**`{base}` に1つへまとめる。**実績として5通りに分かれた"))

    out = root / "out"
    latest = out / LATEST
    if not latest.is_dir():
        findings.append(("out/" + LATEST, "いま見る版が無い",
                         "**`out/latest/` を作る。**見る場所が1つに決まらない"))
    if out.is_dir():
        for d in sorted(x.name for x in out.iterdir() if x.is_dir()):
            if d != LATEST and not re.fullmatch(r"v\d{2,}", d) \
                    and not re.fullmatch(r"latest-[a-z0-9]+", d):
                findings.append((f"out/{d}", "版の名前が決めごとと違う",
                                 "**`vNN`（外に出した版）か `latest`。**"
                                 "2案を並べるときだけ `latest-a` の形"))

    # 日時入りの名前（フォルダもファイルも）
    for p in root.rglob("*"):
        if p.is_dir() and p.name == "work":
            continue
        if _DATE.search(p.name):
            notes.append((str(p.relative_to(root)), "名前に日時が入る（要確認）",
                          "**日付は `project.md` の履歴に書く。**名前では版を区別しない"))

    if findings or notes:
        print("\n| 場所 | 種類 | 中身 |")
        print("|---|---|---|")
        for f in findings + notes:
            print(f"| {f[0]} | {f[1]} | {f[2]} |")
    else:
        print(f"**指摘なし。**見る場所は {latest}")

    code = checked.summary("layout_check", len(dirs), "直下のフォルダ", len(findings), {
        "案件": str(root),
        "いま見る版": str(latest) if latest.is_dir() else "**無い**",
        "要確認": f"{len(notes)} 件",
    })
    return code or (1 if findings else 0)


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""読み手の目の段：枚ごとの「何の枚か」と流れを1つの表にし、**全体像の1文が無い枚**を拾う。

    python3 scripts/story_check.py <検査する>.pptx

見た目・日本語の検査が全部通っても、「いきなり画面の説明に入り、何の形かが分からない」枚は残る
（画面の画像と凡例だけの枚は、検査を通っても、読み手には何の枚かが分からない）。

| 見るもの | 区分 | 理由 |
|---|---|---|
| 画像が枚の 35% 以上を占め、見出し以外の文が 25字未満 | **要確認** | 画面を見せる枚に「何を、どういう形にしたか」の1文が無い |
| 見出しが空 | 要確認 | 何の枚かが分からない |
| 枚ごとの見出しと最初の1文の一覧 | 表（判定しない） | **前後のつながりは人が読む。**この表を上から読んで流れを確かめる |

**この道具は「流れが良い」とは言わない。**表を出し、上の2つを拾うだけ。流れの判断は表を読んだ人がする。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import checked  # noqa: E402

PIC_RATIO = 0.35
LEAD_MIN = 25


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    from pptx import Presentation
    prs = Presentation(Path(sys.argv[1]).expanduser().resolve())
    area = prs.slide_width * prs.slide_height
    rows, notes = [], []
    for no, sl in enumerate(prs.slides, 1):
        title, texts, pic = "", [], 0
        for sh in sl.shapes:
            if sh.shape_type == 13 and sh.width and sh.height:
                pic += sh.width * sh.height
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                t = " ".join(p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip())
                if not t:
                    continue
                if sh.is_placeholder and "TITLE" in str(sh.placeholder_format.type) and not title:
                    title = t
                else:
                    texts.append(t)
        body = "".join(texts)
        lead = texts[0][:40] if texts else "（文が無い）"
        ratio = pic / area if area else 0
        rows.append((no, title or "（見出し無し）", lead, f"{ratio:.0%}"))
        if not title:
            notes.append((no, "見出しが空", "何の枚かが分からない"))
        if ratio >= PIC_RATIO and len(body) < LEAD_MIN:
            notes.append((no, "画面が主で、全体像の1文が無い（推定）",
                          f"画像 {ratio:.0%}・文 {len(body)}字。「何を、どういう形にしたか」を1文で足すか、ノートの冒頭で言うか"))
    print("## 流れの表（上から読み、前の枚とのつながりを確かめる）\n")
    print("| 枚 | 見出し | 最初の1文 | 画像の占める割合 |")
    print("|---|---|---|---|")
    for r in rows:
        print(f"| {r[0]} | {r[1][:30]} | {r[2]} | {r[3]} |")
    if notes:
        print("\n## 要確認\n")
        print("| 枚 | 種類 | 中身 |")
        print("|---|---|---|")
        for n in notes:
            print(f"| {n[0]} | {n[1]} | {n[2]} |")
    return checked.summary("story_check", len(rows), "枚", 0, {"要確認": f"{len(notes)} 件（流れの良し悪しは判定しない）"})


if __name__ == "__main__":
    sys.exit(main())

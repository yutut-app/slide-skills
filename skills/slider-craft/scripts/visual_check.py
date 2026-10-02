#!/usr/bin/env python3
"""図版と余白を機械で検査する。**絵を人に見せる前に必ず通す。**

    python3 scripts/visual_check.py work/*.html --manifest <テンプレの manifest> \
        --review work/visual-review.md

## 何を見るか

| 見るもの | 落ちる条件 |
|---|---|
| 本文の上端・下端 | テンプレの基準値から **40px 超**ずれている（全枚でそろわない） |
| 文字の大きさ | 本文・表・軸・注記が、テンプレの下限 px を下回る |
| 文字どうしの重なり | テキストを持つ枠が重なっている（ラベルが線や他のラベルに乗る） |
| 同じ数値の二重 | 1枚に同じ数値が2回以上出る（**要確認**。右軸とラベルの二重など） |
| 枠に入らない値 | 枠の高さが文字の高さに足りない（**値が黙って消える**） |
| 判定表 | `--review` の表に「未確認」が残っている |

**指摘があれば異常終了する**（終了コード 1）。「要確認」だけなら落とさない。

## 判定を4値で出す

合 / 否 / 要確認 / **未確認**。**既定は未確認。**
見ていないものを「合」と書かない（見ていないことと問題が無いことは別）。

## 基準値はテンプレートごとに持つ

**この台本に数値を埋め込まない。**投影する環境もテンプレートも案件ごとに違う。
`--manifest` に書く。無いときは**その旨を前提欄に出し、上下端と文字サイズは見ない。**

```
body-top: 190
body-bottom: 664
min-font: 13.3
```
"""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked

TOL = 40.0          # 上下端のずれの許容（px）
_NUM = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)\s*(%|％|件|円|人|分|台|回|倍)")


def manifest_of(path: Path):
    """`key: value` を拾う。**無い鍵は見ない**（勝手な既定値を置かない）。"""
    out = {}
    if not path or not path.exists():
        return out
    for line in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^\s*\|?\s*([a-z-]+)\s*[:|]\s*([\d.]+)", line)
        if m:
            out[m.group(1)] = float(m.group(2))
    return out


def boxes(html: Path):
    """[(枚, 要素, 左, 上, 幅, 高, 文字, フォント px)]。**UTF-8 で明示的に読む。**"""
    import lxml.html as LH

    doc = LH.document_fromstring(html.read_text(encoding="utf-8"))
    for bad in doc.xpath("//script|//style"):
        bad.getparent().remove(bad)
    slides = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]") or [doc]
    out = []
    for i, sl in enumerate(slides, 1):
        for el in sl.iter():
            st = (el.get("style") or "").replace(" ", "")
            def px(name):
                m = re.search(name + r":(-?[\d.]+)px", st)
                return float(m.group(1)) if m else None
            left, top, w, h = px("left"), px("top"), px("width"), px("height")
            if left is None or top is None:
                continue
            txt = re.sub(r"\s+", " ", "".join(el.itertext())).strip()
            out.append((i, el, left, top, w, h, txt, px("font-size")))
    return out


def overlap(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    if None in (aw, ah, bw, bh):
        return 0.0
    dx = min(ax + aw, bx + bw) - max(ax, bx)
    dy = min(ay + ah, by + bh) - max(ay, by)
    return dx * dy if dx > 0 and dy > 0 else 0.0


def main():
    ap = argparse.ArgumentParser(description="図版と余白を機械で検査する")
    ap.add_argument("html", nargs="+")
    ap.add_argument("--manifest", help="テンプレの基準値（body-top / body-bottom / min-font）")
    ap.add_argument("--review", help="図版の判定表。**未確認が残っていれば落とす**")
    args = ap.parse_args()

    man = manifest_of(Path(args.manifest)) if args.manifest else {}
    top_ref, bot_ref = man.get("body-top"), man.get("body-bottom")
    min_font = man.get("min-font")

    findings, notes, n_slides = [], [], 0
    for f in args.html:
        path = Path(f)
        rows = boxes(path)
        slides = sorted({r[0] for r in rows})
        n_slides += len(slides)
        for no in slides:
            items = [r for r in rows if r[0] == no]
            body = [r for r in items
                    if top_ref is None or r[3] >= top_ref - TOL]
            # 上端・下端が全枚でそろうか
            if top_ref is not None and body:
                t = min(r[3] for r in body)
                if abs(t - top_ref) > TOL:
                    findings.append((path.name, no, "本文の上端がずれる",
                                     f"{t:.0f}px / 基準 {top_ref:.0f}px"))
            if bot_ref is not None and body:
                b = max((r[3] + (r[5] or 0)) for r in body if r[3] < bot_ref)
                if abs(b - bot_ref) > TOL:
                    findings.append((path.name, no, "本文の下端がずれる",
                                     f"{b:.0f}px / 基準 {bot_ref:.0f}px。**下の空きが揃わない**"))
            # 文字の大きさ
            if min_font is not None:
                for r in items:
                    if r[7] and r[6] and r[7] < min_font - 0.05:
                        findings.append((path.name, no, "文字が小さい",
                                         f"{r[7]:.1f}px / 下限 {min_font:.1f}px「{r[6][:12]}」"))
            # 文字どうしの重なり
            texts = [r for r in items if r[6] and r[4] and r[5]]
            for i in range(len(texts)):
                for j in range(i + 1, len(texts)):
                    a, b = texts[i], texts[j]
                    if b[1] in a[1].iterancestors() or a[1] in b[1].iterancestors():
                        continue
                    if overlap(a[2:6], b[2:6]) > 0:
                        findings.append((path.name, no, "文字が重なる",
                                         f"「{a[6][:10]}」と「{b[6][:10]}」"))
            # 枠に入らない値（**黙って消えるのが最悪**）
            for r in texts:
                if r[5] and r[7] and r[5] < r[7] * 1.1:
                    findings.append((path.name, no, "枠に値が入らない",
                                     f"高さ {r[5]:.0f}px に {r[7]:.1f}px の「{r[6][:8]}」。"
                                     "**外に出す**"))
            # 同じ数値の二重（要確認）
            seen = {}
            for r in texts:
                for m in _NUM.finditer(r[6]):
                    k = m.group(0).replace(" ", "")
                    seen[k] = seen.get(k, 0) + 1
            for k, c in seen.items():
                if c > 1:
                    notes.append((path.name, no, "同じ数値が二重（要確認）",
                                  f"「{k}」が {c} 箇所。**右軸とラベルの二重など**"))

    # 判定表の未確認
    if args.review:
        rv = Path(args.review)
        if not rv.exists():
            findings.append((rv.name, "—", "判定表が無い",
                             "**図版の観点（軸・目盛・原点・単位・ラベル）を枚ごとに出す**"))
        else:
            for i, line in enumerate(rv.read_text(encoding="utf-8").splitlines(), 1):
                if "未確認" in line:
                    findings.append((rv.name, i, "判定表に未確認が残る", line.strip()[:60]))

    if findings or notes:
        print("\n| ファイル | 枚 | 種類 | 中身 |")
        print("|---|---|---|---|")
        for x in findings + notes:
            print(f"| {x[0]} | {x[1]} | {x[2]} | {x[3]} |")
    else:
        print("**指摘なし。**この後、`references/80_audience.md` の図版の観点で目で見る。")

    code = checked.summary("visual_check", n_slides, "枚", len(findings), {
        "基準値": Path(args.manifest).name if man
                  else "**渡されていない（上下端と文字サイズは見ていない）**",
        "判定表": Path(args.review).name if args.review else "**渡されていない**",
        "許容": f"上下端 ±{TOL:.0f}px",
        "要確認": f"{len(notes)} 件",
    })
    return code or (1 if findings else 0)


if __name__ == "__main__":
    sys.exit(main())

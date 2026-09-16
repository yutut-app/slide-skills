#!/usr/bin/env python3
"""資料の HTML から、**資料テンプレート一式を作り直す。**

    python3 scripts/gen_template.py <受け取った>/print.html -o assets/templates/deck/<名前>

## なぜ要るか

**テンプレートは git に入らない。** 別の機械では手元に無い。
無いまま変換すると px→pt が既定（96dpi）に落ち、**スライド寸法が狂う**
（実測で 780×540pt が 960×664pt になった）。

**資料の HTML は、テンプレートと同じ `<style>` を持っている。**
だから**受け取った HTML から作り直せる。**

**使うたびに作り直す。** 手で写すと、配布元がテンプレートを直したときに
古いまま残る。**毎回 HTML から作れば、常に受け取ったものと一致する。**

## 作るもの

| | 中身 |
|---|---|
| `template-manifest.md` | **px→pt の係数**、スライド寸法、共通パーツの座標表、配色、文字サイズ |
| `images/` | ロゴなど。HTML が参照しているものを写す |

**レイアウトの HTML は作らない。** 変換には要らない。
新しい枚を足すなら、資料の HTML をそのまま雛形にする。

## 係数の出どころ

**HTML の `<meta name="px-per-pt">` を見る。** 無ければ `--px-per-pt` で渡す。
**どちらも無いなら作らない。** 既定で作ると、狂った寸法を正しいものとして固定する。

外に出さない。ローカルにファイルを書くだけで、公開もアップロードもしない。
"""

import argparse
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lxml import html as LH

import checked
import html_abs as HA

# 共通パーツとして拾う手がかり。**見た目のためだけの class は拾わない**
SKIP_PREFIX = ("bg-", "text-", "border-", "font-")


def expand(paths):
    out = []
    for s in paths:
        p = Path(s)
        if not p.exists():
            sys.exit(f"見つからない: {p}")
        if p.name == "print.html":
            frames = re.findall(r'<iframe[^>]*src="([^"]+)"',
                                p.read_text(encoding="utf-8"))
            out += [p.parent / f for f in frames]
        else:
            out.append(p)
    return out


def rules_of(files):
    rules = []
    for f in files:
        doc = LH.parse(str(f)).getroot()
        for s in doc.xpath("//style"):
            rules += HA.parse_css(s.text or "")
    return rules


def meta_factor(files):
    for f in files:
        m = re.search(r'<meta\s+name="px-per-pt"\s+content="([\d.]+)"',
                      f.read_text(encoding="utf-8")[:4000])
        if m:
            return float(m.group(1)), f.name
    return None, None


def main():
    ap = argparse.ArgumentParser(
        description="資料の HTML から資料テンプレートを作り直す（ローカルのみ）")
    ap.add_argument("html", nargs="+", help="資料の HTML。print.html でもよい")
    ap.add_argument("-o", "--out", required=True, help="テンプレートの置き場所")
    ap.add_argument("--px-per-pt", type=float,
                    help="px→pt の係数。HTML に meta が無いときに渡す")
    ap.add_argument("--force", action="store_true", help="既にあっても書き換える")
    args = ap.parse_args()

    files = expand(args.html)
    out = Path(args.out)
    man = out / "template-manifest.md"
    if man.exists() and not args.force:
        sys.exit(f"既にある: {man}\n"
                 "  **上書きする前に、どちらが新しいかを確かめる。**\n"
                 "  作り直すなら --force")

    factor, where = args.px_per_pt, "指定"
    if not factor:
        factor, where = meta_factor(files)
        where = f"{where} の meta" if where else None
    if not factor:
        sys.exit("**px→pt の係数が分からない。**\n"
                 "  HTML に <meta name=\"px-per-pt\" content=\"...\"> が無く、\n"
                 "  --px-per-pt も渡されていない。\n"
                 "  **既定で作らない。**狂った寸法を正しいものとして固定することになる。\n"
                 "  配布元に係数を聞く（元 PPTX の pt 寸法 ÷ HTML の px 寸法）。")

    rules = rules_of(files)
    doc0 = LH.parse(str(files[0])).getroot()
    w_px, h_px = HA.slide_size_px(doc0)
    w_pt, h_pt = w_px / factor, h_px / factor

    # 共通パーツ = 位置と大きさが決まっている要素
    parts, colors, sizes = [], {}, {}
    for sel, d in rules:
        sel = sel.strip()
        name = sel.lstrip(".")
        # **座標表に載せるのは、意味のある単独の class だけ。**
        # 色や配置のためだけの class（bg- text- など）は場所の名前にならない
        if name and " " not in name and not name.startswith(SKIP_PREFIX) \
                and all(k in d for k in ("left", "top")):
            parts.append((name, d.get("left", ""), d.get("top", ""),
                          d.get("width", "—"), d.get("height", "—")))
        # **配色と文字サイズは、どのセレクタからも拾う。**
        # 実際の値は `.bg-brand-dark` や `.cover-title p` のような
        # 子孫セレクタ・ユーティリティ側にある。座標表と同じ条件で絞ると
        # **1色・0段階しか出ない**（実際にそうなった）
        for key in ("background-color", "color", "background"):
            c = (d.get(key) or "").strip()
            m = re.match(r"(#[0-9a-fA-F]{3,8})", c)
            if m:
                colors.setdefault(m.group(1).lower(), set()).add(sel)
        fs = d.get("font-size")
        if fs:
            sizes.setdefault(fs, set()).add(sel)

    out.mkdir(parents=True, exist_ok=True)
    lines = [
        "---",
        f"template_name: {out.name}",
        f"slide_size: {w_px:.0f}x{h_px:.0f}px "
        f"(元PPTX {w_pt:.0f}x{h_pt:.0f}pt, 1pt = {factor}px)",
        "generated_from: 資料の HTML（gen_template.py）",
        "---",
        "",
        "# テンプレートマニフェスト",
        "",
        "**資料の HTML から作り直したもの。**",
        f"px→pt の係数の出どころ: {where}",
        "",
        "## 共通パーツの座標",
        "",
        "| 要素 | class | left | top | width | height |",
        "|---|---|---|---|---|---|",
    ]
    for name, l, tp, w, h in sorted(parts):
        lines.append(f"| — | `{name}` | {l} | {tp} | {w} | {h} |")
    lines += ["", "## 配色", "", "| hex | 使う場所 |", "|---|---|"]
    for c, names in sorted(colors.items()):
        lines.append(f"| `{c}` | {', '.join(sorted(names)[:6])} |")
    lines += ["", "## 文字サイズ", "", "| px | pt | 使う場所 |", "|---|---|---|"]
    for fs, names in sorted(sizes.items(), key=lambda kv: -(HA.px(kv[0]) or 0)):
        pt = HA.px(fs)
        lines.append(f"| {fs} | {pt / factor:.1f} | {', '.join(sorted(names)[:6])} |"
                     if pt else f"| {fs} | — | {', '.join(sorted(names)[:6])} |")
    man.write_text("\n".join(lines) + "\n", encoding="utf-8")

    # 画像を写す
    copied = 0
    imgs = set()
    for f in files:
        d = LH.parse(str(f)).getroot()
        imgs |= {s for s in d.xpath("//img/@src")}
    for src in sorted(imgs):
        s = (files[0].parent / src)
        if s.exists():
            dst = out / src
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(s, dst)
            copied += 1

    missing = len(imgs) - copied
    print(f"テンプレートを作った: {out}/\n")
    print("| | |")
    print("|---|---|")
    print(f"| スライド寸法 | {w_px:.0f}×{h_px:.0f}px = {w_pt:.0f}×{h_pt:.0f}pt |")
    print(f"| px→pt の係数 | 1pt = {factor}px（{where}） |")
    print(f"| 共通パーツ | {len(parts)} 件 |")
    print(f"| 配色 | {len(colors)} 色 |")
    print(f"| 文字サイズ | {len(sizes)} 段階 |")
    print(f"| 画像 | {copied} 件{'' if not missing else f'（**{missing} 件が見つからない**）'} |")
    print("\n**係数が合っているかを、変換して確かめる。**")
    print("  出力の「スライドサイズ」が、元 PPTX と同じ pt になっていること。")

    return checked.summary("gen_template", len(parts), "件", missing, {
        "元": f"{len(files)}ファイル",
        "係数": factor,
        "出どころ": where,
    })


if __name__ == "__main__":
    sys.exit(main())

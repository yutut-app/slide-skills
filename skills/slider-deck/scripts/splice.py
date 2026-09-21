#!/usr/bin/env python3
"""テンプレートの HTML の**本文だけ**を入れ替える。共通パーツは触らない。

    python3 scripts/splice.py <テンプレ>/003.html --list
    python3 scripts/splice.py <テンプレ>/003.html --content p05.json -o work/p05.html

## なぜ要るか

**資料作りの本体は「共通パーツを保ったまま本文を入れ替える」こと。**
ヘッダー・フッター・ロゴ・罫は毎枚同じで、変わるのは文言だけ。

**手で書き換えると、共通パーツの座標を動かしてしまう。**
テンプレートの実測値なので、動かすと元の資料と並べたときに合わなくなる。
18枚あれば18回その危険を踏む。**機械でやる。**

**入れ替え忘れも機械で出す。** テンプレートの記入例をそのまま提出する事故は、
検査で最も多い型（`<slider-deck>/references/60_qa.md`）。

## 差し替えの単位

**class 名で指す。** テンプレートは意味のある class（`lead-band`、
`content-title` など）で組んであるので、それをそのまま鍵にする。

```json
{
  "content-title": ["不良の内訳"],
  "lead-band":     ["1行目", "2行目"],
  "note":          {"html": "<table>…</table>"}
}
```

| 書き方 | 意味 |
|---|---|
| `["文字列", ...]` | その要素の `<p>` を、この並びで置き換える |
| `{"html": "…"}` | **中身を丸ごと差し替える**（表・図の入れ物など） |
| `{"skip": true}` | **意図して触らない。**未差し替えの一覧から外す |

**`<p>` の class 属性は元のものを引き継ぐ。** 文字サイズがそこで決まるため。

## 何を触らないか

**位置・大きさ・書体・配色に一切書き込まない。**
このスクリプトは `<p>` の中身と、指定された要素の中身しか変えない。

外に出さない。ローカルにファイルを書くだけで、公開もアップロードもしない。
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lxml import html as LH

import checked

# 共通パーツ。**既定では差し替えの対象から外す。**
# 毎枚同じで、変えるとテンプレートから外れる
COMMON = {"footer-bar", "footer-page", "header-logo", "cover-logo", "logo",
          "header-rule", "divider", "conf-box", "slide"}

# **見た目のためだけの class。**要素の呼び名にしない。
# 色や配置の class は複数の要素が共有するので、鍵にすると
# 「1つ替えたつもりが別の要素も替わる」「一覧に同じ名前が並ぶ」が起きる
UTILITY_PREFIX = ("bg-", "text-", "border-", "font-", "p-", "m-", "w-", "h-")
UTILITY = {"tb", "vmid", "hmid", "center", "right", "left", "bold", "nowrap"}


def is_utility(c):
    return c in UTILITY or c.startswith(UTILITY_PREFIX)


def name_of(el):
    """その要素を指すのに使う class。**意味のあるものを選ぶ。**

    見た目だけの class（色・配置）は複数の要素が共有するので鍵にしない。
    """
    cs = [c for c in classes_of(el) if not is_utility(c) and c not in COMMON]
    return cs[-1] if cs else None


def editable(doc):
    """差し替えの対象になる要素を、**要素単位で**返す。

    class 単位で数えると、色の class を共有する別々の要素が
    同じ行にまとまってしまう（実際に一覧が重複した）。
    """
    out = []
    for el in doc.iter():
        if not isinstance(el.tag, str):
            continue
        if any(c in COMMON for c in classes_of(el)):
            continue
        n = name_of(el)
        if n and el.findall("p") and text_of(el):
            out.append((n, el))
    return out


def classes_of(el):
    return (el.get("class") or "").split()


def text_of(el):
    return " / ".join(t.strip() for t in el.itertext() if t.strip())


def targets(doc):
    """差し替えられる要素を、class 名で引けるようにする。

    **同じ class が複数あれば全部返す。** 1つだけ替えて残りを忘れる事故を防ぐ。
    """
    out = {}
    for el in doc.iter():
        if not isinstance(el.tag, str):
            continue
        for c in classes_of(el):
            out.setdefault(c, []).append(el)
    return out


def replace_paragraphs(el, lines):
    """要素の中の `<p>` を、与えられた行で置き換える。

    **class 属性は引き継ぐ。** そこで文字サイズが決まっているため、
    落とすと見た目が変わる。行数が増えたら最後の `<p>` の class を使う。
    """
    ps = el.findall("p")
    keep = [p.get("class") for p in ps] or [None]
    for p in ps:
        el.remove(p)
    for i, line in enumerate(lines):
        p = LH.Element("p")
        cls = keep[i] if i < len(keep) else keep[-1]
        if cls:
            p.set("class", cls)
        # **\n は <br /> にする。**意味の切れ目の手動改行を殺さない
        parts = str(line).split("\n")
        p.text = parts[0]
        for extra in parts[1:]:
            br = LH.Element("br")
            br.tail = extra
            p.append(br)
        el.append(p)


def replace_html(el, markup):
    for child in list(el):
        el.remove(child)
    el.text = None
    frag = LH.fragment_fromstring(markup, create_parent="div")
    el.text = frag.text
    for child in frag:
        el.append(child)


def main():
    ap = argparse.ArgumentParser(
        description="テンプレートの本文だけを入れ替える（共通パーツは触らない）。枚を足すときに使う")
    ap.add_argument("template", help="元にするテンプレートの HTML")
    ap.add_argument("--content", help="差し替えの中身（JSON）")
    ap.add_argument("-o", "--out", help="出力先の HTML")
    ap.add_argument("--list", action="store_true",
                    help="差し替えられる場所と、いまの文言を出す")
    ap.add_argument("--force", action="store_true", help="出力先を上書きする")
    ap.add_argument("--slide", type=int,
                    help="**何枚目だけを替えるか。**1ファイルに複数の枚がある資料で使う。"
                         "タイトル・リード帯など、全枚にある共通パーツを1枚だけ替えるとき")
    args = ap.parse_args()

    src = Path(args.template)
    if not src.exists():
        sys.exit(f"見つからない: {src}")
    doc = LH.parse(str(src)).getroot()

    # **1ファイルに複数の枚があるなら、どの枚かを決めてから触る。**
    # 共通パーツ（タイトル・リード帯）は全枚に同じ class で入っているので、
    # 枚を絞らずに替えると**全枚が同じ文言になる。**
    slides = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]")
    scope = doc
    if args.slide:
        if not (1 <= args.slide <= len(slides)):
            sys.exit(f"--slide {args.slide} が範囲外（全 {len(slides)} 枚）")
        scope = slides[args.slide - 1]
    found = targets(scope)

    if args.list:
        items = editable(scope)
        counts = {}
        for n, _ in items:
            counts[n] = counts.get(n, 0) + 1
        print(f"差し替えられる場所: {len(items)} 件  （{src.name}）\n")
        print("| class | 同名 | いまの文言 |")
        print("|---|---|---|")
        for n, el in items:
            print(f"| `{n}` | {counts[n]} | {text_of(el)[:46]} |")
        rows = items
        print("\n**共通パーツは出していない**（ヘッダー・フッター・ロゴ・罫）。")
        print("触る必要があるなら、JSON にその class を書けば替えられる。")
        return checked.summary("splice --list", len(rows), "件", 0,
                               {"元": src.name})

    if not args.content or not args.out:
        ap.error("--content と -o を渡す（見るだけなら --list）")

    content = json.loads(Path(args.content).read_text(encoding="utf-8"))
    if len(slides) > 1 and not args.slide:
        hit = [k for k in content if len(found.get(k, [])) > 1]
        if hit:
            sys.exit(f"**{len(slides)} 枚ある資料で、{'・'.join(hit)} が複数の枚にある。**\n"
                     "このまま替えると全枚が同じ文言になる。`--slide N` で枚を指定する")
    out = Path(args.out)
    if out.exists() and not args.force:
        sys.exit(f"既にある: {out}\n  別名で出すか、置き換えるなら --force")

    missing, done, skipped = [], [], set()
    for key, value in content.items():
        els = found.get(key)
        if not els:
            missing.append(key)
            continue
        if isinstance(value, dict) and value.get("skip"):
            skipped.add(key)
            continue
        for el in els:
            if isinstance(value, dict) and "html" in value:
                replace_html(el, value["html"])
            else:
                lines = value if isinstance(value, list) else [value]
                replace_paragraphs(el, lines)
        done.append((key, len(els)))

    # **触っていない場所を出す。**テンプレートの記入例が残る事故がここで出る
    touched = {id(el) for k in content if k in found for el in found[k]}
    untouched = [(n, text_of(el)[:40]) for n, el in editable(doc)
                 if id(el) not in touched and n not in skipped]

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(LH.tostring(doc, encoding="unicode", method="html"),
                   encoding="utf-8")

    print(f"差し替え: {len(done)} 種 → {out}  （元 {src.name}）\n")
    print("| class | 差し替えた数 |")
    print("|---|---|")
    for k, n in done:
        print(f"| `{k}` | {n} |")

    if missing:
        print("\n**テンプレートに無い class を指定している。**"
              "綴りを疑う（`--list` で確かめる）。")
        for k in missing:
            print(f"- {k}")
    if untouched:
        print("\n**触っていない場所に文言が残っている。**"
              "テンプレートの記入例なら、そのまま提出することになる。")
        for c, t in untouched:
            print(f"- `{c}`: {t}")
        print("**意図して残すなら、JSON に `\"<class>\": {\"skip\": true}` と書く。**")

    # **出力先にテンプレートの付属物が無いと、変換が既定値に落ちる。**
    # マニフェストが読めないと px→pt が 96dpi になり、スライド寸法が狂う
    # （実測で 780×540pt が 960×664pt になった）。画像も出なくなる
    lacking = [n for n in ("template-manifest.md", "images")
               if not (out.parent / n).exists()]
    if lacking:
        print(f"\n**出力先に {'/'.join(lacking)} が無い。**")
        print("  マニフェストが無いと px→pt が既定（96dpi）に落ち、"
              "**スライド寸法が狂う。**")
        print("  images が無いとロゴが出ない。")
        print(f"  テンプレートから写す:")
        for n in lacking:
            print(f"    cp -R {src.parent / n} {out.parent}/")
        print("  **変換の出力に出る「出どころ」を必ず読む。**"
              "「既定（96dpi）」なら写せていない。")

    print("\n**位置・大きさ・書体は触っていない。** 溢れるなら文言を減らす。")
    return checked.summary("splice", len(done), "種",
                           len(missing) + len(untouched) + len(lacking),
                           {"元": src.name, "出力": str(out)})


if __name__ == "__main__":
    sys.exit(main())

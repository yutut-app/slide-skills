#!/usr/bin/env python3
"""資料テンプレート（.pptx）から、編集用の HTML を起こす。

    python3 scripts/make_html.py --template <テンプレ>.pptx -o work/deck.html
    python3 scripts/make_html.py --template <テンプレ>.pptx --outline outline.md -o work/deck.html

**HTML が編集対象。** ここで作った HTML を人が直し、区切りがついたら
`html2pptx.py` で pptx にしてレビューする。レビューの指摘は HTML に戻す。

テンプレートのスライドサイズ・書体・テーマ配色を CSS に写すので、
**ブラウザでの見え方が pptx に近くなる。** ただし一致はしない。
最終的な体裁は pptx にしてから `qa_render.py` で確かめる。

--outline は 1行1スライドの素朴な形。

    # 治具交換頻度の変更について        ← 見出し（# の行が新しいスライド）
    > 不良の6割が3号機に集中している     ← メッセージライン
    - 不良率2.4%（目標1.0%）            ← 箇条書き
      - 1.4ポイントの超過               ← 入れ子
    : 2026年1月〜6月、n=600             ← 出典
    // 話す内容                         ← スピーカーノート
"""

import argparse
import html as H
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

EMU = 914400
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def theme(prs):
    try:
        part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    except KeyError:
        return {}, (None, None)
    root = etree.fromstring(part.blob)
    colors = {}
    scheme = root.find(f".//{A}clrScheme")
    if scheme is not None:
        for child in scheme:
            name = child.tag.split("}")[-1]
            srgb = child.find(f"{A}srgbClr")
            sysc = child.find(f"{A}sysClr")
            if srgb is not None:
                colors[name] = "#" + srgb.get("val")
            elif sysc is not None and sysc.get("lastClr"):
                colors[name] = "#" + sysc.get("lastClr")
    ea = latin = None
    mf = root.find(f".//{A}fontScheme/{A}minorFont")
    if mf is not None:
        e, l = mf.find(f"{A}ea"), mf.find(f"{A}latin")
        ea = e.get("typeface") if e is not None else None
        latin = l.get("typeface") if l is not None else None
    return colors, (ea, latin)


def used_fonts(prs):
    """スライドで実際に使われている書体。テーマに ea が無いときの手掛かり。"""
    from collections import Counter
    c = Counter()
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        c[r.font.name] += 1
    return c


def parse_outline(path: Path):
    slides, cur = [], None
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("#"):
            cur = {"title": line.lstrip("#").strip(), "message": "",
                   "bullets": [], "source": "", "notes": ""}
            slides.append(cur)
        elif cur is None:
            continue
        elif line.startswith(">"):
            cur["message"] = line.lstrip(">").strip()
        elif line.lstrip().startswith("-"):
            indent = len(line) - len(line.lstrip())
            cur["bullets"].append((indent // 2, line.lstrip()[1:].strip()))
        elif line.startswith(":"):
            cur["source"] = line.lstrip(":").strip()
        elif line.startswith("//"):
            cur["notes"] = line.lstrip("/").strip()
    return slides


def bullets_html(items):
    if not items:
        return ""
    out, depth = [], 0
    for lvl, text in items:
        while depth < lvl:
            out.append("<ul>"); depth += 1
        while depth > lvl:
            out.append("</ul></li>"); depth -= 1
        out.append(f"<li>{H.escape(text)}")
        out.append("</li>")
    while depth > 0:
        out.append("</ul></li>"); depth -= 1
    return "<ul>" + "".join(out) + "</ul>"


def slide_html(s, layout):
    parts = [f'<section class="slide" data-layout="{layout}">']
    if s.get("title"):
        parts.append(f"  <h1>{H.escape(s['title'])}</h1>")
    if s.get("subtitle"):
        parts.append(f'  <p class="subtitle">{H.escape(s["subtitle"])}</p>')
    if s.get("message"):
        parts.append(f'  <p class="message">{H.escape(s["message"])}</p>')
    if s.get("bullets"):
        parts.append("  " + bullets_html(s["bullets"]))
    if s.get("source"):
        parts.append(f'  <p class="source">{H.escape(s["source"])}</p>')
    if s.get("notes"):
        parts.append(f'  <aside class="notes">{H.escape(s["notes"])}</aside>')
    parts.append("</section>")
    return "\n".join(parts)


SAMPLE = [
    ({"title": "（資料の件名）", "subtitle": "（部署名）　（年月）"}, "title"),
    ({"title": "（見出し・体言止め）", "message": "（メッセージライン・言い切り・数値を入れる）",
      "bullets": [(0, "（箇条書き・体言止め）"), (1, "（入れ子は階層になる）")],
      "notes": "（話す内容。スライドには出ない）"}, "message-body"),
    ({"title": "（見出し）", "message": "（メッセージライン）",
      "source": "（期間・件数・出所）"}, "table"),
]


def main():
    ap = argparse.ArgumentParser(description="テンプレートから編集用の HTML を起こす")
    ap.add_argument("--template", "-t", required=True)
    ap.add_argument("--outline", help="1行1要素の素案（省くと雛形を3枚出す）")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args()

    tpl = Path(args.template).resolve()
    if not tpl.exists():
        sys.exit(f"テンプレートが見つからない: {tpl}")
    out = Path(args.out)

    prs = Presentation(str(tpl))
    w = prs.slide_width / EMU
    h = prs.slide_height / EMU
    colors, (ea, latin) = theme(prs)
    uf = used_fonts(prs)
    # テーマに日本語が無いことがある。そのときは実際に使われている書体を採る
    font_ja = ea or (uf.most_common(1)[0][0] if uf else None) or "sans-serif"
    font_en = latin or "sans-serif"
    accent = colors.get("accent1", "#1F4E79")
    ink = colors.get("dk1", "#333333")
    pale = colors.get("lt2", "#F2F2F2")

    if args.outline:
        data = [(s, "message-body" if s["bullets"] else "section")
                for s in parse_outline(Path(args.outline))]
        if data:
            data[0] = (data[0][0], "title")
    else:
        data = SAMPLE

    try:
        rel = tpl.relative_to(out.resolve().parent)
        tpl_ref = str(rel)
    except ValueError:
        tpl_ref = str(tpl)

    body = "\n\n".join(slide_html(s, lay) for s, lay in data)
    doc = f'''<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<!-- 変換先のテンプレート。html2pptx.py がここを読む -->
<meta name="template" content="{H.escape(tpl_ref)}">
<title>{H.escape(out.stem)}</title>
<style>
/* テンプレート {H.escape(tpl.name)} の体裁を写したもの。
   **見え方の目安であって、pptx と一致はしない。**
   最終確認は pptx にしてから qa_render.py で行う。
   ここの CSS を変えても pptx には反映されない。位置は data-layout が決める。 */
:root {{
  --slide-w: {w:.3f}in; --slide-h: {h:.3f}in;
  --accent: {accent}; --ink: {ink}; --pale: {pale};
  --font-ja: "{H.escape(font_ja)}"; --font-en: "{H.escape(font_en)}";
}}
body {{ background:#666; margin:0; padding:24px;
       font-family: var(--font-en), var(--font-ja), sans-serif; color: var(--ink); }}
.slide {{ width: var(--slide-w); height: var(--slide-h); background:#fff;
         margin:0 auto 24px; padding: 0.5in 0.6in; box-sizing:border-box;
         position:relative; box-shadow:0 2px 8px rgba(0,0,0,.4); overflow:hidden; }}
.slide h1 {{ font-size:24pt; color: var(--accent); margin:0 0 .12in; font-weight:600; }}
.slide .message {{ font-size:13pt; font-weight:700; margin:0 0 .22in; }}
.slide .subtitle {{ font-size:14pt; color:#666; }}
.slide ul {{ font-size:14pt; line-height:1.5; margin:0; padding-left:1.2em; }}
.slide ul ul {{ font-size:12pt; }}
.slide table {{ border-collapse:collapse; width:100%; font-size:11pt; }}
.slide th {{ background:var(--accent); color:#fff; text-align:left; padding:.06in .1in; }}
.slide td {{ background:var(--pale); padding:.06in .1in; }}
.slide .source {{ position:absolute; left:.6in; bottom:.4in; font-size:9pt; color:#777; }}
.slide .notes {{ display:none; }}  /* ノートはスライドに出さない */
[data-layout="title"] {{ display:flex; flex-direction:column; justify-content:center; }}
</style>
</head>
<body>

{body}

</body>
</html>
'''
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(doc, encoding="utf-8")

    print(f"作成: {out}")
    print(f"テンプレート: {tpl.name}  {w:.2f}×{h:.2f}in  "
          f"書体: 日本語={font_ja} / 欧文={font_en}")
    if not ea and uf:
        print("  テーマに日本語（ea）が無いので、実際に使われている書体から決めた")
    print(f"スライド: {len(data)} 枚")
    print("\n**これが編集対象。** ブラウザで開いて直す。")
    print("区切りがついたら pptx にしてレビューする。")
    print(f"  python3 <slider-deck>/scripts/html2pptx.py {out} -o {out.with_suffix('.pptx')}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""手で直した pptx と、HTML テンプレートの差を出す。

    python3 scripts/pptx_diff_html.py <直したもの>.pptx <テンプレ>/print.html

**利用者が PowerPoint 上で直した内容を、HTML テンプレートに戻すために使う。**
pptx が正になるのは、この差分を取る一度きり。戻したら HTML が正に戻る。

出す差分は3つ。

| 種類 | 意味 |
|---|---|
| 文字 | 文言が変わった。HTML の該当箇所を書き換える |
| 位置・大きさ | 図形を動かした・伸ばした。CSS の left/top/width/height を直す |
| 書体・文字サイズ | font-size / 色 / 太字を変えた。CSS を直す |
| 行間 | 段落の行間を変えた。CSS の line-height を直す |

**自動では HTML を書き換えない。** どれを採るかは人が決める。
CSS の1箇所が複数のスライドに効くため、機械的に当てると他が崩れる。
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lxml import html as LH
from pptx import Presentation
from pptx.util import Emu

import checked
import html_abs as HA

# **定数を import 時に控えない。** 換算係数はテンプレートごとに違い、
# apply_scale_from() で差し替わる。控えると 96dpi のまま比較して全項目ずれる。
TOL_PX = 1.0        # これ未満の差は丸め誤差とみなす
TOL_PT = 0.3


def lead_of_pptx(p):
    """段落の行間を (種別, 値) で返す。("x", 倍率) か ("pt", ポイント)。"""
    v = p.line_spacing
    if v is None:
        return None
    return ("pt", v.pt) if hasattr(v, "pt") else ("x", float(v))


def lead_of_html(style, font_px):
    """CSS の line-height を pptx と同じ形に直す。"""
    n = HA.unitless(style.get("line-height"))
    if n is not None:
        return ("x", n)
    v = HA.px(style.get("line-height"))
    if v is not None:
        return ("pt", v / HA.PX_PER_PT)
    return ("x", 1.0)          # 指定なし＝単一


def lead_text(lead):
    if lead is None:
        return "—"
    kind, v = lead
    return f"{v:.2f}倍" if kind == "x" else f"{v:.1f}pt"


def lead_differs(a, b):
    """種別が違えば pt に揃えられないので、そのまま差とみなす。"""
    if a is None or b is None:
        return False
    if a[0] != b[0]:
        return True
    return abs(a[1] - b[1]) >= (0.02 if a[0] == "x" else 0.3)


def default_pt(prs):
    """run に sz が無いとき実際に効く既定サイズ。

    **これを埋めないと、利用者が「既定に戻す」操作をした箇所を見落とす。**
    PowerPoint 上ではサイズが変わって見えるのに、run には属性が無いため。
    """
    el = prs._element
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    hit = el.findall(
        ".//{http://schemas.openxmlformats.org/presentationml/2006/main}"
        "defaultTextStyle/{%s}lvl1pPr/{%s}defRPr" % (ns["a"], ns["a"]))
    for d in hit:
        if d.get("sz"):
            return int(d.get("sz")) / 100
    return 18.0


def shapes_of(slide, dflt=None):
    out = []
    for sh in slide.shapes:
        try:
            if None in (sh.left, sh.top, sh.width, sh.height):
                continue
        except (AttributeError, ValueError):
            continue
        text = sh.text_frame.text.strip() if sh.has_text_frame else ""
        size, lead = None, None
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        size = r.font.size.pt
                        break
                if size:
                    break
            if size is None and text:
                size = dflt
            for p in sh.text_frame.paragraphs:
                lead = lead_of_pptx(p)
                if lead is not None:
                    break
            if lead is None:
                lead = ("x", 1.0)     # 指定なし＝単一
        out.append({
            "text": text,
            "left": sh.left / HA.EMU_PER_PX, "top": sh.top / HA.EMU_PER_PX,
            "width": sh.width / HA.EMU_PER_PX, "height": sh.height / HA.EMU_PER_PX,
            "pt": size, "lead": lead,
        })
    return out


def html_slide_count(path: Path) -> int:
    """その HTML に `.slide` がいくつあるか。**1ファイル＝1枚とは限らない。**"""
    doc = LH.parse(str(path)).getroot()
    return len(doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]"))


def html_boxes(path: Path, index: int = 0):
    """HTML の `index` 枚目から、pptx に出るはずの図形を同じ形で取り出す。

    **`.slide` が複数あれば、その何枚目かを指定して読む。**
    先頭だけを見ていたため、18枚を1ファイルにした資料を「1枚」と数え、
    **残り17枚を見ないまま「差は無い」と報告した。**
    """
    doc = LH.parse(str(path)).getroot()
    rules = []
    for s in doc.xpath("//style"):
        rules += HA.parse_css(s.text or "")
    slides = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]")
    if not slides or index >= len(slides):
        return []
    body = doc.find(".//body")
    inherit = HA.inheritable(HA.computed(body, rules) if body is not None else {})
    out = []
    for el in slides[index].iterchildren():
        if not isinstance(el.tag, str):
            continue
        st = HA.computed(el, rules, inherit)
        left, top = HA.px(st.get("left")), HA.px(st.get("top"))
        if left is None or top is None:
            continue
        blocks = HA.text_blocks(el, rules, HA.inheritable(st))
        text = "\n".join(b[1] for b in blocks)
        pt, lead = None, None
        if blocks:
            fs = HA.px(blocks[0][0].get("font-size"))
            pt = round(fs / HA.PX_PER_PT, 1) if fs else None
            lead = lead_of_html(blocks[0][0], fs)
        out.append({
            "text": text.strip(),
            "left": left, "top": top,
            "width": HA.px(st.get("width")), "height": HA.px(st.get("height")),
            "pt": pt, "lead": lead,
            "class": el.get("class") or el.tag,
        })
    return out


def pair(html_items, pptx_items):
    """同じ図形どうしを対応づける。まず文字で、次に位置の近さで。"""
    pairs, used = [], set()
    for hi, h in enumerate(html_items):
        best, score = None, None
        for pi, p in enumerate(pptx_items):
            if pi in used:
                continue
            if h["text"] and h["text"] == p["text"]:
                best, score = pi, -1
                break
            d = abs(h["left"] - p["left"]) + abs(h["top"] - p["top"])
            if score is None or d < score:
                best, score = pi, d
        if best is not None and (score == -1 or score < 40):
            used.add(best)
            pairs.append((h, pptx_items[best]))
        else:
            pairs.append((h, None))
    extra = [p for i, p in enumerate(pptx_items) if i not in used]
    return pairs, extra


def compare(html_file: Path, pptx_slide, no, dflt=None, index: int = 0):
    rows = []
    h_items = html_boxes(html_file, index)
    p_items = shapes_of(pptx_slide, dflt)
    pairs, extra = pair(h_items, p_items)

    for h, p in pairs:
        who = f"`.{h['class'].split()[-1]}`" if h.get("class") else "?"
        if p is None:
            rows.append((no, who, "消えた", "—", "pptx 側に無い"))
            continue
        if h["text"] != p["text"]:
            rows.append((no, who, "文字",
                         (h["text"] or "（空）").replace("\n", " / ")[:26],
                         (p["text"] or "（空）").replace("\n", " / ")[:26]))
        for k, label in (("left", "left"), ("top", "top"),
                         ("width", "width"), ("height", "height")):
            hv, pv = h.get(k), p.get(k)
            if hv is None or pv is None:
                continue
            if abs(hv - pv) >= TOL_PX:
                rows.append((no, who, f"位置 {label}",
                             f"{hv:.1f}px", f"{pv:.1f}px"))
        if h["pt"] and p["pt"] and abs(h["pt"] - p["pt"]) >= TOL_PT:
            rows.append((no, who, "文字サイズ",
                         f"{h['pt']:.1f}pt", f"{p['pt']:.1f}pt"))
        if lead_differs(h.get("lead"), p.get("lead")):
            rows.append((no, who, "行間",
                         lead_text(h["lead"]), lead_text(p["lead"])))
    for p in extra:
        rows.append((no, "—", "増えた", "—",
                     (p["text"] or "（図形）").replace("\n", " / ")[:26]))
    return rows


def main():
    ap = argparse.ArgumentParser(
        description="手で直した pptx と HTML テンプレートの差を出す。手で直された pptx を受け取ったときに使う")
    ap.add_argument("pptx")
    ap.add_argument("html", nargs="+",
                    help="スライド順の HTML。print.html を渡すと iframe 順に展開する")
    args = ap.parse_args()

    pptx_path = Path(args.pptx)
    if not pptx_path.exists():
        sys.exit(f"見つからない: {pptx_path}")

    # **単位はファイルではなく「枚」。**(ファイル, その中の何枚目) で持つ
    files = []
    for h in args.html:
        p = Path(h)
        if not p.exists():
            sys.exit(f"見つからない: {p}")
        doc = LH.parse(str(p)).getroot()
        frames = doc.xpath("//iframe/@src")
        for f in ([p.parent / x for x in frames] if frames else [p]):
            n = html_slide_count(f)
            files += [(f, i) for i in range(max(1, n))]

    scale, src = HA.apply_scale_from([f for f, _ in files])
    print(f"px→pt 換算: 1pt = {scale}px（出どころ: {src}）")

    prs = Presentation(str(pptx_path))
    slides = list(prs.slides)
    print(f"pptx: {pptx_path.name}  {len(slides)} 枚")
    names = []
    for f, idx in files:
        names.append(f.name if idx == 0 and
                     sum(1 for g, _ in files if g == f) == 1
                     else f"{f.name}#{idx + 1}")
    print(f"HTML: {len(files)} 枚  （{', '.join(names)}）")
    if len(slides) != len(files):
        # **数が合わないまま比べない。**足りない方に合わせて黙って打ち切ると、
        # 見ていない枚があるのに「差は無い」と報告することになる
        print(f"\n**枚数が違う（pptx {len(slides)} / HTML {len(files)}）。**")
        print("  スライドの追加・削除があったか、**渡す HTML が足りない。**")
        print("  **このまま比べない。**少ない方に合わせると、"
              "見ていない枚を「差は無い」と報告することになる。")
        return checked.summary("pptx_diff_html", 0, "枚", 1,
                               {"pptx": len(slides), "HTML": len(files)})

    rows = []
    for i, ((f, idx), s) in enumerate(zip(files, slides), 1):
        rows += compare(f, s, i, default_pt(prs), idx)

    if not rows:
        print("\n差は無い。HTML テンプレートは pptx と一致している。")
        return 0

    print(f"\n## 差分 {len(rows)} 件\n")
    print("| # | 要素 | 種類 | HTML（今） | pptx（直した後） |")
    print("|---|---|---|---|---|")
    for r in rows:
        print("| " + " | ".join(str(x) for x in r) + " |")
    print("\n**採る差分を選んでから HTML を直す。自動では書き換えない。**")
    print("CSS の1箇所が複数のスライドに効くので、機械的に当てると他が崩れる。")
    print("直したら再変換して、`diff_slides.py` で意図した箇所だけが変わったか確かめる。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

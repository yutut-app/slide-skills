#!/usr/bin/env python3
"""px 絶対座標の HTML を pptx にする。`html2pptx.py` から使う。

**テンプレートが HTML 一式（マニフェスト方式）のときの経路。**
テンプレートの .pptx を土台にせず、**HTML に書かれた座標をそのまま pptx に写す。**

    1px = 9525 EMU（96px/inch）。1280x720px = 13.333x7.5in = 16:9。
    元 PPTX が 960x540pt なら 1pt = 1.3333px で完全に一致する。

**値を変えずにテキストだけ差し替えれば、元 PPTX と同じ見た目に戻る。**
これはマニフェスト方式のテンプレートが持つ性質で、この経路はそれを壊さない。

## 対応している CSS

必要なものだけを見る。**CSS 全般を解釈しない。**

| 見るもの | 使い道 |
|---|---|
| `position/left/top/width/height` | 図形の位置と大きさ |
| `font-size/font-weight/color/font-family` | 文字（継承する） |
| `text-align/line-height` | 段落 |
| `background-color` | 塗り |
| `border` | 枠線 |
| `padding` | テキストの内側余白（PowerPoint のインセット） |
| `display:flex` ＋ `justify-content:center` | 上下中央ぞろえ |

セレクタは `.class`、`tag`、`.a .b`（子孫）、`.a.b`（複合）、カンマ区切りのみ。
擬似クラス（`:first-child` など）とメディアクエリは**無視する**。
"""

import re
import sys
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Pt

EMU_PER_PX = 9525          # 96px/inch → 914400/96
PX_PER_PT = 4 / 3          # 1pt = 1.3333px（既定。テンプレートによって違う）
EMU_PER_PT = 12700


def template_manifests(files):
    """手元のテンプレート置き場にある `template-manifest.md` を集める。

    探す順は **環境変数 → HTML の上の階層 → 作業ディレクトリ**。
    README には「別の場所に置くなら `SLIDE_TEMPLATE_DIR` を向ける」と
    書いてあったが、**このスキルには探す仕組みが無かった。**
    書いてあるだけの決まりは守られない。
    """
    import os

    roots = []
    env = os.environ.get("SLIDE_TEMPLATE_DIR")
    if env:
        roots.append(Path(env))
    bases = [Path(f).resolve().parent for f in files] + [Path.cwd().resolve()]
    for b in bases:
        for up in [b] + list(b.parents)[:4]:
            roots.append(up / "assets" / "templates" / "deck")
    found, seen = [], set()
    for r in roots:
        if not r.is_dir():
            continue
        for man in sorted(r.glob("*/template-manifest.md")):
            key = man.resolve()
            if key not in seen:
                seen.add(key)
                found.append(man)
        direct = r / "template-manifest.md"
        if direct.exists() and direct.resolve() not in seen:
            seen.add(direct.resolve())
            found.append(direct)
    return found


def px_per_pt_of(files, override=None):
    """px→pt の換算係数を決める。

    **96dpi 決め打ちにしない。** HTML を起こした人が元 PPTX の pt 座標を
    どの係数で px にしたかはテンプレートごとに違う。
    同じ場所の `template-manifest.md` に `1pt = <n>px` があればそれを使う。

    **この係数を使うものはすべてこの関数を通す。** 変換だけ直して差分を
    直さないと、差分が全項目ずれたまま「差がある」と出る（実際に起きた）。
    """
    if override:
        return float(override), "指定"
    for d in sorted({Path(f).resolve().parent for f in files}):
        man = d / "template-manifest.md"
        if not man.exists():
            continue
        m = re.search(r"1\s*pt\s*=\s*([\d.]+)\s*px", man.read_text())
        if m:
            return float(m.group(1)), f"{man.parent.name}/template-manifest.md"
    # **HTML 自身が持っていれば、それを使う。**
    # テンプレートは git に入らないので、別の機械では手元に無いことがある。
    # 係数が読めないと 96dpi に落ち、スライド寸法が狂う（実測で 780→960pt）。
    # 資料の HTML に書いておけば、テンプレートが無くても正しく変換できる
    for f in files:
        try:
            head = Path(f).read_text(encoding="utf-8")[:4000]
        except OSError:
            continue
        m = re.search(r'<meta\s+name="px-per-pt"\s+content="([\d.]+)"', head)
        if m:
            return float(m.group(1)), f"{Path(f).name} の meta"

    # **手元のテンプレートを見に行く。**
    # 受け取った HTML と同じ場所にマニフェストは無い（テンプレートは別置き）。
    # 探さないと既定の 96dpi に落ち、テンプレートによっては寸法が狂う
    # （A4 は 1pt = 1.641px。実測で 780×540pt が 960×664pt になった）。
    mans = template_manifests(files)
    if len(mans) == 1:
        m = re.search(r"1\s*pt\s*=\s*([\d.]+)\s*px", mans[0].read_text())
        if m:
            return float(m.group(1)), f"手元のテンプレ {mans[0].parent.name}"
    elif len(mans) > 1:
        print("**手元にテンプレートが複数あり、どれか決められない**"
              f"（{'・'.join(sorted(x.parent.name for x in mans))}）。"
              "`--px-per-pt` で渡すか、HTML に meta を書く", file=sys.stderr)
    return 4 / 3, "既定（96dpi）"


def apply_scale_from(files, override=None):
    """files の場所からマニフェストを探して換算係数を適用する。戻り値は (係数, 出どころ)。"""
    scale, src = px_per_pt_of(files, override)
    set_scale(scale)
    return scale, src


def set_scale(px_per_pt: float):
    """px→pt の換算係数を差し替える。

    **テンプレートごとに違う。** 96dpi（1pt=1.3333px）とは限らない。
    HTML を起こした人が、元 PPTX の pt 座標を好きな係数で px にしているため。

    実績: 16:9 版は 1.333333（1280px=960pt）だが、A4 版は 1.641026
    （1280px=780pt）。既定のまま A4 版を変換すると、スライドが
    960×664pt になり、元 PPTX の 780×540pt と一致しない。
    """
    global PX_PER_PT, EMU_PER_PX
    if not px_per_pt or px_per_pt <= 0:
        raise ValueError(f"px_per_pt が不正: {px_per_pt}")
    PX_PER_PT = float(px_per_pt)
    EMU_PER_PX = EMU_PER_PT / PX_PER_PT
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

INHERITED = ("font-size", "font-weight", "color", "font-family",
             "line-height", "text-align", "white-space")


# ---------------------------------------------------------------------------
# ごく小さな CSS エンジン
# ---------------------------------------------------------------------------

def parse_css(text):
    """`セレクタ { 宣言 }` を並びにする。@ルールとコメントは捨てる。"""
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    rules = []
    for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", text):
        sels = [s.strip() for s in m.group(1).split(",") if s.strip()]
        decls = {}
        for d in m.group(2).split(";"):
            if ":" not in d:
                continue
            k, v = d.split(":", 1)
            decls[k.strip().lower()] = v.strip()
        for sel in sels:
            if sel.startswith("@") or "%" in sel:
                continue
            rules.append((sel, decls))
    return rules


def _match_simple(el, part):
    """`.a`, `.a.b`, `tag`, `tag.a`, `*` を判定する。"""
    if part == "*":
        return True
    part = re.sub(r":[a-zA-Z-]+(\([^)]*\))?", "", part)  # 擬似クラスは落とす
    if not part:
        return True
    tag = re.match(r"^[a-zA-Z][\w-]*", part)
    if tag:
        if el.tag != tag.group(0):
            return False
        part = part[len(tag.group(0)):]
    classes = set((el.get("class") or "").split())
    for cls in re.findall(r"\.([\w-]+)", part):
        if cls not in classes:
            return False
    return True


def matches(el, selector):
    """子孫結合子つきのセレクタに合うか。右から左へ確かめる。"""
    parts = selector.split()
    if not parts:
        return False
    if not _match_simple(el, parts[-1]):
        return False
    cur = el.getparent()
    for part in reversed(parts[:-1]):
        while cur is not None and not _match_simple(cur, part):
            cur = cur.getparent()
        if cur is None:
            return False
        cur = cur.getparent()
    return True


def specificity(selector):
    return (len(re.findall(r"\.", selector)), len(selector.split()))


def computed(el, rules, inherit=None):
    """要素の実効スタイル。詳細度の低い順に重ね、最後に inline を当てる。"""
    out = dict(inherit or {})
    hits = [(specificity(sel), i, d) for i, (sel, d) in enumerate(rules)
            if matches(el, sel)]
    for _, _, decls in sorted(hits, key=lambda x: (x[0], x[1])):
        out.update(decls)
    inline = el.get("style")
    if inline:
        for d in inline.split(";"):
            if ":" in d:
                k, v = d.split(":", 1)
                out[k.strip().lower()] = v.strip()
    return out


def inheritable(style):
    return {k: v for k, v in style.items() if k in INHERITED}


# ---------------------------------------------------------------------------
# 値の読み取り
# ---------------------------------------------------------------------------

def px(value, default=None):
    """`12px` のように単位が px のときだけ数値を返す。

    **単位なしの数値を px とみなさない。** line-height の `1.35` は倍率であり、
    px として扱うと行高が 1pt 未満になって文字が消える（実際に消えた）。
    """
    if value is None:
        return default
    s = str(value).strip()
    if re.fullmatch(r"-?0+(\.0+)?", s):
        return 0.0          # CSS ではゼロだけ単位を省ける（left: 0 など）
    m = re.search(r"(-?[\d.]+)\s*px", s)
    return float(m.group(1)) if m else default


def unitless(value):
    """単位なしの数値（line-height の倍率など）。px 付きなら None。"""
    s = str(value).strip() if value is not None else ""
    if value is None or "px" in s or re.fullmatch(r"-?0+(\.0+)?", s):
        return None
    m = re.fullmatch(r"\s*(-?[\d.]+)\s*", str(value))
    return float(m.group(1)) if m else None


def color(value):
    if not value:
        return None
    v = value.strip()
    m = re.fullmatch(r"#([0-9a-fA-F]{6})", v)
    if m:
        return RGBColor.from_string(m.group(1).upper())
    m = re.fullmatch(r"#([0-9a-fA-F]{3})", v)
    if m:
        s = "".join(c * 2 for c in m.group(1))
        return RGBColor.from_string(s.upper())
    m = re.match(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", v)
    if m:
        return RGBColor(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return {"white": RGBColor(0xFF, 0xFF, 0xFF),
            "black": RGBColor(0, 0, 0)}.get(v.lower())


def padding_px(style):
    """`padding: 上下 左右` / 個別指定 をまとめて返す。"""
    top = right = bottom = left = 0.0
    if "padding" in style:
        vals = [px(v, 0.0) for v in style["padding"].split()]
        if len(vals) == 1:
            top = right = bottom = left = vals[0]
        elif len(vals) == 2:
            top = bottom = vals[0]; right = left = vals[1]
        elif len(vals) == 3:
            top, right, bottom = vals; left = right
        elif len(vals) >= 4:
            top, right, bottom, left = vals[:4]
    for side, name in ((("top",), "padding-top"), (("right",), "padding-right"),
                       (("bottom",), "padding-bottom"), (("left",), "padding-left")):
        if name in style:
            v = px(style[name], 0.0)
            if side[0] == "top":
                top = v
            elif side[0] == "right":
                right = v
            elif side[0] == "bottom":
                bottom = v
            else:
                left = v
    return top, right, bottom, left


def font_faces(style):
    """font-family から、欧文と日本語の書体を1つずつ取り出す。

    'Verdana', 'Meiryo', sans-serif → 欧文=Verdana / 和文=Meiryo。
    **和文は latin でなく ea に書く必要がある。**
    """
    fam = style.get("font-family", "")
    names = [n.strip().strip("'\"") for n in fam.split(",") if n.strip()]
    names = [n for n in names if n.lower() not in
             ("sans-serif", "serif", "monospace", "system-ui")]
    if not names:
        return None, None
    latin = names[0]
    KNOWN_JA = ("meiryo", "meiryo ui", "yu gothic", "ms pgothic", "ms gothic",
                "hiragino sans", "noto sans jp", "biz udpgothic", "biz udgothic")
    cands = [n for n in names
             if re.search(r"[ぁ-んァ-ヶ一-龯]", n) or n.lower() in KNOWN_JA]
    # 日本語表記の別名（メイリオ）があればそちらを採る。
    # 元 PPTX の表記に合わせておくと、PowerPoint のフォント欄の表示が一致する
    ja = next((n for n in cands if re.search(r"[ぁ-んァ-ヶ一-龯]", n)), None)
    return latin, ja or (cands[0] if cands else latin)


def set_theme_fonts(prs, latin, ja):
    """テーマの見出し用・本文用フォントを差し替える。

    **これをしないと、PowerPoint のフォント欄には既定（Calibri / ＭＳ Ｐゴシック）が
    出て、新しく打った文字がその書体になる。** 既存の文字は run に直接書いてあるので
    正しく見えるため、テンプレートとして配るまで気づけない。
    """
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    except (KeyError, IndexError):
        return False
    root = etree.fromstring(part.blob)
    scheme = root.find(f".//{A}fontScheme")
    if scheme is None:
        return False
    for kind in ("majorFont", "minorFont"):
        node = scheme.find(f"{A}{kind}")
        if node is None:
            continue
        el = node.find(f"{A}latin")
        if el is not None and latin:
            el.set("typeface", latin)
        ea = node.find(f"{A}ea")
        if ea is None:
            ea = node.makeelement(f"{A}ea", {})
            node.insert(1, ea)
        if ja:
            ea.set("typeface", ja)
        # script 別フォント（Jpan）も合わせる。ここが ＭＳ Ｐゴシックのままだと
        # 日本語だけ別の書体になる
        for f in node.findall(f"{A}font"):
            if f.get("script") == "Jpan" and ja:
                f.set("typeface", ja)
    part._blob = etree.tostring(root, xml_declaration=True,
                                encoding="UTF-8", standalone=True)
    return True


# ---------------------------------------------------------------------------
# pptx を組む
# ---------------------------------------------------------------------------

def set_run(run, style, base_ja=None, base_latin=None, force_ja=False):
    size = px(style.get("font-size"))
    if size:
        run.font.size = Pt(round(size / PX_PER_PT, 1))
    if style.get("font-weight") in ("bold", "700", "800", "900"):
        run.font.bold = True
    c = color(style.get("color"))
    if c is not None:
        run.font.color.rgb = c
    latin, ja = font_faces(style)
    latin = latin or base_latin
    ja = ja or base_ja
    # **和文だけの run は latin にも和文書体を書く。**
    # 描画は変わらないが、PowerPoint のフォント欄に和文書体が出る
    if force_ja and ja:
        latin = ja
    if latin:
        run.font.name = latin
    rPr = run.font._rPr
    for tag, face in (("ea", ja), ("cs", latin)):
        if not face:
            continue
        el = rPr.find(f"{A}{tag}")
        if el is None:
            el = rPr.makeelement(f"{A}{tag}", {})
            rPr.append(el)
        el.set("typeface", face)


# **和文として扱う文字。**かな・漢字・全角記号に加えて、
# 矢印・囲み数字・幾何学記号・文字様記号も入れる。
# これらは全角幅で組まれるので、欧文書体に渡すと字幅が変わり、
# 前後の字間が崩れる（→ ① ◎ ○ △ ■ ▲ ℃ で実際に起きた）
_JA_RE = re.compile(
    r"[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u3400-\u4dbf\u4e00-\u9fff"
    r"\uf900-\ufaff\uff00-\uff60\uffe0-\uffe6"
    r"\u2010-\u201f\u2030-\u205e\u2100-\u214f\u2190-\u21ff\u2460-\u24ff"
    r"\u25a0-\u25ff\u2600-\u26ff\u3200-\u32ff]")


def split_ja(text):
    """文字列を、和文の塊と欧文の塊に切る。

    **1つの run に和文と欧文を混ぜると、run には書体を1つしか書けない。**
    latin に欧文・ea に和文を入れて PowerPoint に任せることになり、
    描画は正しいが**フォント欄には latin しか出ず**、
    「和文まで欧文書体になっている」と読めてしまう。

    切れば和文 run の latin にも和文書体を書けるので、
    **見た目は同じまま、フォント欄の表示も一致する。**

    空白は直前の塊に付ける。**切り口をむやみに増やさない。**
    """
    out, cur, cur_ja = [], "", None
    for ch in text:
        ja = bool(_JA_RE.match(ch))
        if ch.isspace() and cur:
            cur += ch
            continue
        if cur_ja is None or ja == cur_ja:
            cur += ch
            cur_ja = ja
        else:
            out.append((cur_ja, cur))
            cur, cur_ja = ch, ja
    if cur:
        out.append((bool(cur_ja), cur))
    return out


def add_runs(p, text, st, ja0, latin0):
    """段落に、和文／欧文で切り分けた run を並べる。"""
    for is_ja, chunk in split_ja(text):
        r = p.add_run()
        r.text = chunk
        set_run(r, st, base_ja=ja0, base_latin=latin0, force_ja=is_ja)


ALIGN = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
         "left": PP_ALIGN.LEFT, "justify": PP_ALIGN.JUSTIFY}


def text_blocks(el, rules, inherit):
    """要素の中から、段落になるものを (スタイル, 文字) で返す。

    <p> / <li> を段落にする。無ければ要素自身の文字を1段落にする。
    <br> は段落の区切りにする。
    """
    blocks = []
    for node in el.iter():
        if node is el or node.tag not in ("p", "li"):
            continue
        st = computed(node, rules, inherit)
        # **大きさ・色・太さが、<p> ではなく中の <span> に書かれていることがある。**
        # <p> だけを見ると既定の文字サイズに落ち、折り返しが変わって枠からこぼれる。
        # 中身が1つの要素だけに包まれているときは、そちらの体裁を使う
        inner = [k for k in node if isinstance(k.tag, str) and k.tag != "br"]
        if len(inner) == 1 and not (node.text or "").strip() \
                and not (inner[0].tail or "").strip():
            st = computed(inner[0], rules, st)
        # <br> で分ける
        chunks, cur = [], (node.text or "")
        for child in node:
            if child.tag == "br":
                chunks.append(cur)
                cur = child.tail or ""
            else:
                cur += (child.text_content() if hasattr(child, "text_content")
                        else (child.text or "")) + (child.tail or "")
        chunks.append(cur)
        for ch in chunks:
            t = " ".join(ch.split())
            if t:
                blocks.append((st, t))
    if not blocks:
        t = " ".join(el.text_content().split())
        if t:
            blocks.append((computed(el, rules, inherit), t))
    return blocks


def looks_like_chart(el, rules, inherit):
    """**箱で描いたグラフ**（棒を `<div>` で作ったもの）を見つける。

    `<svg data-chart>` はネイティブのグラフになるが、**箱で描いたものは
    ただの塗り図形として写る。**絵は似ていても、中に数字が入っていない。
    受け取った側は PowerPoint でも Excel でも直せない。

    黙って写すと気づけないので、**警告だけ出す。**出力は変えない。
    条件は「塗られた兄弟が3つ以上あり、**大きさがそろっていない**」。
    大きさがそろっているもの（KPI の札など）は、グラフではないので外す。
    """
    kids = [k for k in el.iter() if isinstance(k.tag, str) and k is not el]
    by_parent = {}
    for k in kids:
        by_parent.setdefault(k.getparent(), []).append(k)
    for group in by_parent.values():
        filled = []
        for k in group:
            st = computed(k, rules, inherit)
            if background_color(st) is None:
                continue
            hh, ww = px(st.get("height")), px(st.get("width"))
            if hh is None and ww is None:
                continue
            filled.append((hh, ww))
        if len(filled) < 3:
            continue
        for i in (0, 1):
            vals = [v[i] for v in filled if v[i] is not None]
            if len(vals) >= 3 and len(set(round(v, 1) for v in vals)) >= 3:
                return True
    return False


def shape_kind(style, w_px, h_px):
    """CSS の `border-radius` から、pptx の図形の種類を決める。

    **四角で作ると、HTML と見た目が変わる。**実案件では、角丸69・丸12・
    大きな角丸8 の計89個が**すべて四角になっていた。**
    丸バッジや帯の印象がまるごと変わるので、見た目の検査の前に効く。

    戻り値は (種類, 角丸の深さ)。深さは短い辺に対する割合（0〜0.5）。
    """
    from pptx.enum.shapes import MSO_SHAPE

    raw = str(style.get("border-radius") or "").strip()
    if not raw:
        return MSO_SHAPE.RECTANGLE, None
    short = max(1.0, min(w_px or 1.0, h_px or 1.0))
    if raw.endswith("%"):
        try:
            pct = float(raw.rstrip("%"))
        except ValueError:
            pct = 0.0
        if pct >= 50:
            return MSO_SHAPE.OVAL, None
        r = short * pct / 100
    else:
        r = px(raw) or 0.0
    if r <= 0:
        return MSO_SHAPE.RECTANGLE, None
    # 半径が短い辺の半分以上なら、**両端が半円の「丸帯」**
    return MSO_SHAPE.ROUNDED_RECTANGLE, min(0.5, r / short)


def add_shape(slide, el, style, rules, inherit, base_dir: Path, warn):
    left, top = px(style.get("left")), px(style.get("top"))
    w, h = px(style.get("width")), px(style.get("height"))
    if left is None or top is None:
        return
    if w is None and el.tag in ("svg", "canvas"):
        w = px(el.get("width")) or (float(el.get("width"))
                                    if (el.get("width") or "").isdigit() else None)
    if h is None and el.tag in ("svg", "canvas"):
        h = px(el.get("height")) or (float(el.get("height"))
                                     if (el.get("height") or "").isdigit() else None)
    if w is None:
        w = 100.0
    x, y = Emu(int(left * EMU_PER_PX)), Emu(int(top * EMU_PER_PX))
    cx = Emu(int(w * EMU_PER_PX))

    # 画像
    if el.tag == "img":
        src = el.get("src") or ""
        p = Path(src)
        p = p if p.is_absolute() else (base_dir / p)
        if not p.exists():
            warn.append(f"画像が無い: {src}")
            return
        cy = Emu(int((h or w) * EMU_PER_PX))
        slide.shapes.add_picture(str(p), x, y, cx, cy)
        return

    # 表は**ネイティブの表**にする。画像やテキストの寄せ集めにしない。
    # 受け取った側が数字を直せることが、資料を渡す目的そのもの
    tbl = el if el.tag == "table" else next(iter(el.xpath(".//table")), None)
    if tbl is not None:
        add_table(slide, tbl, style, rules, inherit, x, y, cx, h, warn)
        return

    # **図は入れ子にあることが多い。**表と同じく、中から拾う。
    # 実案件の HTML は本文枠（body-area）の中に図を置く作りで、
    # 直下しか見ていなかったため**グラフが1つも入らず、警告も出なかった**
    if el.tag not in ("svg", "canvas") and not el.get("data-chart"):
        # 数のあるグラフ → ネイティブのグラフ、線の図 → コネクタ。
        # **どちらも入れ子にある。**実案件では 144 個の作図 svg のうち
        # 直下にあった1個しか変換されず、**143 個が黙って落ちた**
        figs = el.xpath(".//*[@data-chart]") + el.xpath(".//svg")
        for fig in figs:
            fst = computed(fig, rules, inherit)
            fl, ft = px(fst.get("left")), px(fst.get("top"))
            fw, fh = px(fst.get("width")), px(fst.get("height"))
            if fw is None and (fig.get("width") or "").replace(".", "").isdigit():
                fw = float(fig.get("width"))
            if fh is None and (fig.get("height") or "").replace(".", "").isdigit():
                fh = float(fig.get("height"))
            fx = left if fl is None else fl
            fy = top if ft is None else ft
            fwidth = w if fw is None else fw
            fheight = (h or w) if fh is None else fh
            if fig.get("data-chart"):
                if add_chart(slide, fig,
                             Emu(int(fx * EMU_PER_PX)), Emu(int(fy * EMU_PER_PX)),
                             Emu(int(fwidth * EMU_PER_PX)),
                             Emu(int(fheight * EMU_PER_PX)), warn):
                    continue
            try:
                n = draw_svg_lines(slide, fig, fx, fy, fwidth, fheight)
                if n:
                    warn.append(f"入れ子の svg の線 {n} 本をコネクタにした")
            except SvgOutside as ex:
                warn.append(f"**入れ子の svg の線がスライドの外へ出た（{ex}）。**"
                            "枠の寸法と `viewBox` が噛み合っていない")

    # svg の扱いは**中身で分かれる。**
    #
    #   数値を表すグラフ → 写さない。**絵から元の数字は戻せない。**
    #                      Excel のテンプレートか、ネイティブのグラフで作る
    #   線と文字で作る図 → **ネイティブのコネクタに写す。**
    #                      写した結果は編集できる図形であって、画像ではない
    if el.tag in ("svg", "canvas") or el.get("data-chart"):
        want = el.get("data-figure") or (el.findtext("{*}title") or "").strip()
        # 数が書いてあれば、**絵を写さずネイティブのグラフにする**
        cy_ = Emu(int((h or w) * EMU_PER_PX))
        if add_chart(slide, el, x, y, cx, cy_, warn):
            return
        if want in CHART_ROUTE:
            warn.append(
                f"**{want} は写さない。**絵から元の数字は戻せない。"
                f" → {CHART_ROUTE[want]}（データを直せる形にする）")
            return
        try:
            n = draw_svg_lines(slide, el, left, top, w, h)
        except SvgOutside as ex:
            warn.append(
                f"**svg の線がスライドの外へ出た（{ex}）。**"
                "枠の `left/top/width/height` と `viewBox` が噛み合っていない。"
                "**viewBox を書くか、CSS で width と height を px で指定する**")
            return
        if n:
            warn.append(f"svg の線 {n} 本をコネクタにした"
                        + (f"（{want}）" if want else "")
                        + "。**編集できる図形として入っている**")
        else:
            warn.append(
                f"**{el.tag} に写せる線が無い**"
                + (f"（{want}）" if want else "")
                + "。写せるのは `<line>` と `<polyline>` だけ。"
                " グラフは `assets/templates/charts/` `qc7/`、"
                "作図は `qc7/` `n7/` を使う（`assets/templates/INDEX.md`）")
        return

    if el.tag not in ("svg", "canvas") and looks_like_chart(el, rules, inherit):
        # **絵は似ていても、中に数字が無い。**受け取った側が直せない
        warn.append(
            "**箱で描いたグラフを、塗り図形のまま写した。**"
            "中に数字が入らないので、PowerPoint でも Excel でも直せない。"
            "配布元に `data-chart` と `data-values` を足してもらう"
            "（足りていれば、中にワークシートを持つグラフになる）")

    blocks = text_blocks(el, rules, inherit)
    fill = color(background_color(style))
    border = style.get("border")
    if h is None:
        # 高さの指定が無い要素は、中身の行数から見積もる。
        # 決め打ちにすると、住所のような複数行の枠が1行分になって溢れる
        total = 0.0
        for st, _ in blocks:
            fs = px(st.get("font-size"), 16.0)
            lh = px(st.get("line-height"))
            if lh is None:
                mult = unitless(st.get("line-height"))
                lh = (mult or 1.35) * fs
            total += lh
        pt_, _, pb_, _ = padding_px(style)
        h = max(total + pt_ + pb_, 16.0) if blocks else 24.0
    cy = Emu(int(h * EMU_PER_PX))

    if not blocks and fill is None and not border:
        return  # 中身も色も無いなら作らない

    if fill is not None or border:
        kind, adj = shape_kind(style, w, h)
        shp = slide.shapes.add_shape(kind, x, y, cx, cy)
        if adj is not None:
            try:
                shp.adjustments[0] = adj
            except (IndexError, ValueError):
                pass       # 角丸を持たない図形。**形は合っているので進める**
        shp.shadow.inherit = False
        st = shp._element.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}style")
        if st is not None:
            shp._element.remove(st)
        if fill is not None:
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
        else:
            shp.fill.background()
        bc = color(re.sub(r"[\d.]+px|solid|dashed|dotted", "", border or "").strip()) \
            if border else None
        if bc is not None:
            shp.line.color.rgb = bc
            bw = px(border, 1.0)
            shp.line.width = Pt(max(0.5, (bw or 1.0) / PX_PER_PT))
        else:
            shp.line.fill.background()
    else:
        shp = slide.shapes.add_textbox(x, y, cx, cy)

    tf = shp.text_frame
    # **`white-space:nowrap` は枠ではなく中の <p> に書かれていることがある。**
    # 枠だけを見ると折り返しありと判定し、丸バッジの文字が2行になって
    # 枠からこぼれ、下の行と重なった
    nowrap = style.get("white-space") == "nowrap" or (
        bool(blocks) and all(st.get("white-space") == "nowrap"
                             for st, _ in blocks))
    tf.word_wrap = not nowrap
    # **枠を自動で広げない。**
    # python-pptx の add_textbox は <a:spAutoFit/> を付けるため、
    # PowerPoint が枠を縦に広げる。HTML は固定高なので、
    # 折り返した枠だけが下へ伸びて隣や下の要素と重なる（18枚中12枚で発生）
    tf.auto_size = MSO_AUTO_SIZE.NONE
    pt_, pr_, pb_, pl_ = padding_px(style)
    tf.margin_left = Emu(int(pl_ * EMU_PER_PX))
    tf.margin_right = Emu(int(pr_ * EMU_PER_PX))
    tf.margin_top = Emu(int(pt_ * EMU_PER_PX))
    tf.margin_bottom = Emu(int(pb_ * EMU_PER_PX))
    # **縦位置は必ず明示する。** 既定に任せない。
    # 塗りのある要素は autoshape になり、python-pptx の既定で中央寄せになる。
    # HTML が上寄せでも pptx は中央、というずれが黙って入る（実際に起きた）。
    # flex は**主軸がどちらか**で意味が変わる。
    # 既定（row）なら justify-content が横、align-items が縦。
    # column なら逆。両方とも縦に倒していたため、
    # `.footer-page{display:flex;justify-content:center}` のページ番号が
    # 横に中央ぞろえされず、箱からはみ出していた。
    # **`display:flex` は、枠ではなく中の要素に書かれていることがある。**
    # 外側だけを見ると中央ぞろえを見落とし、**高い枠ほど上寄せが目立つ**。
    # 体裁の読み取り（text_blocks）と同じで、1つだけの子を辿る
    flex_st = style
    if style.get("display") != "flex":
        node = el
        for _ in range(3):
            kids = [k for k in node if isinstance(k.tag, str) and k.tag != "br"]
            if len(kids) != 1:
                break
            node = kids[0]
            st_in = computed(node, rules, style)
            if st_in.get("display") == "flex":
                flex_st = st_in
                break
    v_center = h_center = False
    if flex_st.get("display") == "flex":
        col = "column" in (flex_st.get("flex-direction") or "")
        jc = flex_st.get("justify-content") == "center"
        ai = flex_st.get("align-items") == "center"
        v_center, h_center = (jc, ai) if col else (ai, jc)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if v_center else MSO_ANCHOR.TOP

    latin0, ja0 = font_faces(inherit)
    for i, (st, txt) in enumerate(blocks):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        align = st.get("text-align", style.get("text-align", ""))
        p.alignment = ALIGN.get(align) or (PP_ALIGN.CENTER if h_center else None)
        lh = st.get("line-height") or style.get("line-height")
        if lh:
            if px(lh) is not None:
                p.line_spacing = Pt(round(px(lh) / PX_PER_PT, 1))
            elif unitless(lh) is not None:
                p.line_spacing = unitless(lh)   # 倍率はそのまま渡す
        add_runs(p, txt, st, ja0, latin0)


# HTML の `data-figure` から、使うテンプレートを引く。
# **図は変換せず、作り直す。**受け取った側が数字を直せる形にするため。
# 対応の全体は `assets/templates/INDEX.md`。ここは名前が一致したときの近道。
FIGURE_ROUTE = {
    "パレート図": "`qc7/01_パレート図.xlsx`",
    "ヒストグラム": "`qc7/02_ヒストグラム.xlsx`",
    "管理図": "`qc7/03_X-R管理図.xlsx`",
    "X-R管理図": "`qc7/03_X-R管理図.xlsx`",
    "散布図": "`charts/08_散布図.xlsx`（層別するなら `qc7/04_層別散布図.xlsx`）",
    "特性要因図": "`qc7/07_特性要因図.pptx`",
    "縦棒グラフ": "`charts/01_縦棒グラフ.xlsx`",
    "横棒グラフ": "`charts/02_横棒グラフ.xlsx`",
    "折れ線グラフ": "`charts/03_折れ線グラフ.xlsx`",
    "積み上げ棒グラフ": "`charts/04_積み上げ棒グラフ.xlsx`",
    "帯グラフ": "`charts/05_帯グラフ.xlsx`",
    "円グラフ": "`charts/06_円グラフ.xlsx`",
    "複合グラフ": "`charts/07_複合グラフ.xlsx`",
    "レーダーチャート": "`charts/09_レーダーチャート.xlsx`",
    "ウォーターフォール": "`charts/10_ウォーターフォール.xlsx`",
    "面グラフ": "`charts/11_面グラフ.xlsx`",
    "親和図": "`n7/01_親和図法.pptx`",
    "連関図": "`n7/02_連関図法.pptx`",
    "系統図": "`n7/03_系統図法.pptx`",
    "マトリックス図": "`n7/04_マトリックス図法.pptx`",
    "アローダイアグラム": "`n7/05_アローダイアグラム法.pptx`",
    "PDPC": "`n7/06_PDPC法.pptx`",
}


# 絵から元の数字が戻せない図。**写さずに、データを持つ形で作り直す。**
CHART_ROUTE = {k: v for k, v in FIGURE_ROUTE.items() if v.endswith(".xlsx`")}


# `data-chart` の値 → pptx のグラフ種別。**図の名前ではなく、型で決める。**
# 「パレート図」のような用途の名前で分けると、他の案件で増え続ける。
CHART_TYPES = {
    "column": "COLUMN_CLUSTERED", "bar": "BAR_CLUSTERED",
    "column-stacked": "COLUMN_STACKED", "bar-stacked": "BAR_STACKED",
    "column-stacked-100": "COLUMN_STACKED_100",
    "bar-stacked-100": "BAR_STACKED_100",
    "line": "LINE_MARKERS", "line-plain": "LINE",
    "pie": "PIE", "doughnut": "DOUGHNUT",
    "area": "AREA", "area-stacked": "AREA_STACKED",
    "radar": "RADAR_MARKERS", "scatter": "XY_SCATTER",
}


def numbers(text):
    """`data-values="62,18,11"` を数の列にする。空なら空。"""
    out = []
    for tok in re.split(r"[,\s]+", (text or "").strip()):
        if not tok:
            continue
        try:
            out.append(float(tok))
        except ValueError:
            return []
    return out


def add_chart(slide, el, x, y, cx, cy, warn):
    """`data-chart` が書いてある要素を、pptx のネイティブのグラフにする。

    **絵を写さない。数を渡す。**
    pptx のグラフは中にワークシートを持つので、受け取った側が
    **PowerPoint の中で数字を直せる。**別に Excel を添える必要がない。

    ```html
    <svg data-chart="column"
         data-labels="A,B,C"
         data-values="62,18,11"
         data-series="不良件数">…</svg>
    ```

    系列を増やすときは `data-values-2` `data-series-2` と続ける。
    **戻り値は、作れたかどうか。**
    """
    kind = (el.get("data-chart") or "").strip().lower()
    if kind not in CHART_TYPES:
        return False
    labels = [t.strip() for t in (el.get("data-labels") or "").split(",") if t.strip()]
    series = []
    for i in range(1, 9):
        suffix = "" if i == 1 else f"-{i}"
        vals = numbers(el.get("data-values" + suffix))
        if not vals:
            continue
        name = (el.get("data-series" + suffix) or f"系列{i}").strip()
        series.append((name, vals))
    if not series:
        warn.append(f"**data-chart=\"{kind}\" に data-values が無い。**"
                    "グラフを作れない。数を書く")
        return False
    if not labels:
        labels = [str(i + 1) for i in range(len(series[0][1]))]

    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    data = CategoryChartData()
    data.categories = labels
    for name, vals in series:
        # 系列の長さは分類にそろえる。**足りない分を勝手に埋めない**
        if len(vals) != len(labels):
            warn.append(f"**系列「{name}」の数が分類と合わない**"
                        f"（値{len(vals)} / 分類{len(labels)}）。短い方に切る")
        n = min(len(vals), len(labels))
        data.add_series(name, vals[:n])

    slide.shapes.add_chart(getattr(XL_CHART_TYPE, CHART_TYPES[kind]),
                           x, y, cx, cy, data)
    extra = numbers(el.get("data-line"))
    if extra:
        warn.append("**`data-line` は第2軸の折れ線で、pptx に入れられない。**"
                    "累積線が要るなら `qc7/01_パレート図.xlsx` で作る")
    warn.append(f"グラフにした（{kind} / 系列{len(series)} / 分類{len(labels)}）。"
                "**PowerPoint の中で数字を直せる**")
    return True


def draw_svg_lines(slide, svg, left, top, w, h):
    """svg の `<line>` `<polyline>` を pptx の直線コネクタにする。

    **推測が入らない。** 座標をそのまま写すだけ。
    viewBox があれば、そこから枠のサイズへ拡大縮小する。

    | svg の属性 | pptx |
    |---|---|
    | `stroke` | 線色 |
    | `stroke-width` | 太さ |
    | `stroke-dasharray` | 破線 |
    | `marker-end` | 終端の矢印 |

    特性要因図の骨、系統図の接続、アローダイアグラムの矢線は、
    **これが無いと文字が浮いているだけになる。**
    """
    from pptx.enum.shapes import MSO_CONNECTOR

    # **viewBox が無ければ拡大縮小しない。** 座標はそのまま px とみなす。
    # 以前は枠の寸法から倍率を作っていたが、寸法が取れないと
    # でたらめな倍率になり、線がスライドの外へ飛んでいた
    # （height が取れないと y だけ 100倍・800倍になる。実際に起きた）。
    # **HTML パーサは属性名を小文字にする。**`viewBox` は `viewbox` で入る。
    # 大文字のまま探していたため、拡大縮小が一度も効いていなかった。
    # 枠と viewBox が 1:1 の入力では正しく見えるので、気づけなかった
    raw_vb = svg.get("viewBox") or svg.get("viewbox") or ""
    vb = raw_vb.replace(",", " ").split()
    vx = vy = 0.0
    sx = sy = 1.0
    if len(vb) == 4:
        try:
            vx, vy, vw, vh = (float(v) for v in vb)
        except ValueError:
            vw = vh = 0.0
        # **枠の寸法が取れたときだけ**倍率を出す。取れなければ等倍
        if vw > 0 and w:
            sx = w / vw
        if vh > 0 and h:
            sy = h / vh

    def emu(px_x, px_y):
        return (Emu(int((left + (px_x - vx) * sx) * EMU_PER_PX)),
                Emu(int((top + (px_y - vy) * sy) * EMU_PER_PX)))

    segments = []
    for e in svg.iter():
        tag = e.tag.split("}")[-1] if isinstance(e.tag, str) else ""
        if tag == "line":
            try:
                segments.append((e, [(float(e.get("x1", 0)), float(e.get("y1", 0))),
                                     (float(e.get("x2", 0)), float(e.get("y2", 0)))]))
            except ValueError:
                continue
        elif tag == "polyline":
            nums = [float(v) for v in
                    re.findall(r"-?[\d.]+", e.get("points", ""))]
            pts = list(zip(nums[0::2], nums[1::2]))
            if len(pts) >= 2:
                segments.append((e, pts))

    drawn = 0
    outside = 0
    sw = slide.part.package.presentation_part.presentation.slide_width / EMU_PER_PX
    sh = slide.part.package.presentation_part.presentation.slide_height / EMU_PER_PX
    for e, pts in segments:
        stroke = e.get("stroke") or style_attr(e, "stroke") or "#000000"
        if stroke.lower() in ("none", "transparent"):
            continue
        width_px = e.get("stroke-width") or style_attr(e, "stroke-width") or "1"
        dashed = bool(e.get("stroke-dasharray") or style_attr(e, "stroke-dasharray"))
        arrow = bool(e.get("marker-end"))
        for i in range(len(pts) - 1):
            (x1, y1), (x2, y2) = pts[i], pts[i + 1]
            a, b = emu(x1, y1), emu(x2, y2)
            cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                            a[0], a[1], b[0], b[1])
            c = color(stroke)
            if c is not None:
                cn.line.color.rgb = c
            try:
                cn.line.width = Pt(max(0.5, float(re.sub(r"[^\d.]", "", width_px) or 1)
                                       / PX_PER_PT))
            except ValueError:
                pass
            ln = cn.line._get_or_add_ln()
            if dashed:
                d = ln.makeelement(f"{A}prstDash", {"val": "dash"})
                ln.append(d)
            # 矢印は最後の区間にだけ付ける
            if arrow and i == len(pts) - 2:
                ln.append(ln.makeelement(f"{A}tailEnd", {"type": "triangle"}))
            drawn += 1
            # **数えるだけでは足りない。**座標が正しいかまで見る。
            # 「8本入った」と報告しながら、全部スライド外だったことがある
            for (px_, py_) in ((x1, y1), (x2, y2)):
                gx = left + (px_ - vx) * sx
                gy = top + (py_ - vy) * sy
                if not (-1 <= gx <= sw + 1 and -1 <= gy <= sh + 1):
                    outside += 1
                    break
    if outside:
        raise SvgOutside(outside, drawn)
    return drawn


class SvgOutside(Exception):
    """線がスライドの外へ出た。**黙って通さない。**"""

    def __init__(self, outside, drawn):
        super().__init__(f"{drawn}本中 {outside}本がスライドの外")
        self.outside, self.drawn = outside, drawn


def style_attr(el, name):
    """`style="stroke:#000"` の形で書かれた値を拾う。"""
    m = re.search(rf"(?:^|;)\s*{name}\s*:\s*([^;]+)", el.get("style") or "")
    return m.group(1).strip() if m else None


def add_table(slide, tbl, style, rules, inherit, x, y, cx, h, warn):
    """`<table>` を pptx のネイティブの表にする。

    **画像やテキストの寄せ集めにしない。** 受け取った側が数字を直せることが、
    資料を渡す目的そのもの（`references/52_charts.md`）。
    **塗りは CSS を写す。** ヘッダの帯色・行の交互塗り・強調行の色は、
    テンプレートが意図して `<td>` に書いている。既定に任せると
    **全部が白や薄いグレーに落ちて、意図が消える。**

    **罫線だけは PowerPoint の既定に任せる。**
    CSS で作った罫線の色や太さは写さない（**配布元の規約**（<slider-craft>/references/48_html.md）「効かないこと」）。
    """
    rows = tbl.xpath(".//tr")
    if not rows:
        return
    ncol = max(len(r.xpath("./td|./th")) for r in rows)
    if ncol == 0:
        return
    cy = Emu(int((h or len(rows) * 28) * EMU_PER_PX))
    shape = slide.shapes.add_table(len(rows), ncol, x, y, cx, cy)
    table = shape.table

    # 列幅は `<colgroup>` にあればそれを使う。**均等割りにしない。**
    # 項目名の列と数値の列が同じ幅になると、折り返して行が崩れる
    widths = []
    for col in tbl.xpath(".//col"):
        wpx = px(col.get("width")) or px(
            (re.search(r"width\s*:\s*([^;]+)", col.get("style") or "") or
             [None, None])[1] if "width" in (col.get("style") or "") else None)
        widths.append(wpx)
    if len(widths) == ncol and all(w_ for w_ in widths):
        total = sum(widths)
        box = cx / EMU_PER_PX
        for ci, w_ in enumerate(widths):
            table.columns[ci].width = Emu(int(w_ / total * box * EMU_PER_PX))

    # 書体と文字サイズは本文と同じ扱いにする。**表だけ既定に戻さない**
    latin0, ja0 = font_faces(inherit)
    for ri, tr in enumerate(rows):
        cells = tr.xpath("./td|./th")
        for ci in range(ncol):
            cell = table.cell(ri, ci)
            txt = ""
            if ci < len(cells):
                txt = " ".join(cells[ci].itertext()).strip()
            cell.text = txt
            cell_style = computed(cells[ci], rules, inherit) if ci < len(cells) \
                else dict(inherit)
            # **セルの中身は td > div > p > span と包まれ、
            # font-size は一番内側にあることがある。**
            # td だけを見ると既定の文字サイズに落ち、表が本文より大きくなって
            # 行が伸び、表がスライドの外へはみ出す
            # **余白は td 自身から取る。**padding は継承されないので、
            # 内側へ降りた後の体裁から読むと 0 になる（実際にそうなった）
            td_style = cell_style
            if ci < len(cells):
                node = cells[ci]
                while True:
                    kids = [k for k in node if isinstance(k.tag, str)]
                    if len(kids) != 1:
                        break
                    node = kids[0]
                    cell_style = computed(node, rules, cell_style)

            # **セルの余白も CSS から取る。**
            # PowerPoint の既定は左右 0.1in（9.6px）で、CSS の padding より広い。
            # 既定のままだと収まるはずの文字が折り返し、行が伸びる
            cpt, cpr, cpb, cpl = padding_px(td_style)
            cell.margin_left = Emu(int(cpl * EMU_PER_PX))
            cell.margin_right = Emu(int(cpr * EMU_PER_PX))
            cell.margin_top = Emu(int(cpt * EMU_PER_PX))
            cell.margin_bottom = Emu(int(cpb * EMU_PER_PX))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

            # **塗りは CSS から写す。**既定に任せると意図が消える
            bg = background_color(td_style)
            fill_rgb = color(bg) if bg else None
            if fill_rgb is not None:
                cell.fill.solid()
                cell.fill.fore_color.rgb = fill_rgb

            # `cell.text = txt` が作った run を捨てて、和文で切り分けて入れ直す
            para0 = cell.text_frame.paragraphs[0]
            for r in list(para0.runs):
                r._r.getparent().remove(r._r)
            add_runs(para0, txt, cell_style, ja0, latin0)

    if any(r.xpath("./th") for r in rows):
        table.first_row = True
    warn.append(f"表を pptx のネイティブの表にした（{len(rows)}行×{ncol}列）。"
                "**塗りは CSS を写した。罫線は PowerPoint の既定に従う。**")


def background_color(style):
    """塗りの色。**`background-color` だけを見ない。**

    `background: #008096` のショートハンドで書かれた塗りが全部消えていた。
    **HTML では正常に見えるので、pptx を見るまで気づけない。**
    ショートハンドの中から色だけを拾う（画像・グラデーションは扱わない）。
    """
    c = style.get("background-color")
    if c:
        return c
    bg = (style.get("background") or "").strip()
    if not bg or "gradient" in bg or "url(" in bg:
        return None
    m = re.search(r"(#[0-9a-fA-F]{3,8}|rgba?\([^)]*\)|\b[a-z]{3,20}\b)", bg)
    if not m:
        return None
    word = m.group(0).lower()
    if word in ("none", "transparent", "inherit", "initial", "unset",
                "repeat", "no-repeat", "center", "cover", "contain", "fixed"):
        return None
    return m.group(0)


def build_slide(prs, doc, base_dir: Path, warn, relayout=False):
    rules = []
    for style_el in doc.xpath("//style"):
        rules += parse_css(style_el.text or "")

    slides = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]")
    if not slides:
        return None

    body = doc.find(".//body")
    base_style = computed(body, rules) if body is not None else {}
    inherit = inheritable(base_style)

    # **1つの HTML に複数の .slide があれば、全部スライドにする。**
    # 以前は先頭だけを変換し、残りを黙って捨てていた。
    # 15枚入りの deck.html を渡して「変換: 1 枚」で正常終了した実績がある。
    made = None
    sw, sh = slide_size_px(doc)
    for root in slides:
        made = prs.slides.add_slide(prs.slide_layouts[6])   # 白紙
        items = []
        for el in root.iterchildren():
            if not isinstance(el.tag, str):
                continue
            st = computed(el, rules, inherit)
            # **本文枠は入れ物であって、図形ではない。**
            # 中に座標を持つ子が複数あるなら、その子ごとに置く。
            # まとめて1つの枠にすると、表と図が同居したときに
            # **表だけが拾われて図が落ちる**（実案件で起きた）
            kids = [(k, computed(k, rules, inheritable(st)))
                    for k in el.iterchildren() if isinstance(k.tag, str)]
            placed = [(k, ks) for k, ks in kids if px(ks.get("left")) is not None]
            if len(placed) >= 2 and not (el.text or "").strip():
                items += placed
            else:
                items.append((el, st))
        if relayout:
            # **座標を写さず、行に束ねて敷き直す。**
            # 写した座標は、PowerPoint が枠を縦に伸ばしたぶんだけ重なる
            import relayout as RL
            RL.relayout(items, sw, sh, warn)
        for el, st in items:
            add_shape(made, el, st, rules, inheritable(st), base_dir, warn)
    return made


def slide_size_px(doc):
    rules = []
    for style_el in doc.xpath("//style"):
        rules += parse_css(style_el.text or "")
    for sel, d in rules:
        if re.search(r"\.slide(?![\w-])", sel) and "width" in d and "height" in d:
            return px(d["width"], 1280.0), px(d["height"], 720.0)
    return 1280.0, 720.0


def looks_absolute(doc):
    """この HTML が px 絶対座標の作りかどうか。"""
    rules = []
    for style_el in doc.xpath("//style"):
        rules += parse_css(style_el.text or "")
    has_slide_px = any(re.search(r"\.slide(?![\w-])", sel) and "width" in d
                       and "px" in str(d.get("width", ""))
                       for sel, d in rules)
    has_abs = any(d.get("position") == "absolute" for _, d in rules)
    return has_slide_px and has_abs

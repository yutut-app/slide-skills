#!/usr/bin/env python3
"""汎用の資料テンプレート（.pptx）を assets/templates/deck/ に作る。

**利用者のテンプレートが1つも無いときの土台。**
出るのは `.pptx` なので、消費者は `make_html.py --template <これ>.pptx`
（このスキルの経路B）。**これだけが使い道。**

**`slider-reskin` の一覧には出ない。**
向こうが探すのは `template-manifest.md` を持つ HTML のディレクトリで、
これが置くのは同じ場所に平置きした `.pptx` ファイル。
**別テンプレートへの載せ替えには使えない。**

**生成物は git に入らない**（`.gitignore`）。消しても作り直せる。

**スライドマスターとレイアウトの構成は既定のまま変えない。**
変えるのはテーマのフォントと配色、スライドサイズだけ。
レイアウトを自作すると、利用者が置くテンプレートとの互換が取れなくなる。

    python3 scripts/gen_deck_templates.py [出力先]
"""

import os
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def template_dir():
    """資料テンプレートの置き場所を探す。

    複数のスキルが同じフォルダを見る必要があるため、
    リポジトリ直下の assets/templates/deck を正とし、単独で置かれた場合に備えて
    スキル内も見る。**テンプレートを2箇所に置かない。**置くと必ず食い違う。
    """
    if os.environ.get("SLIDE_TEMPLATE_DIR"):
        return Path(os.environ["SLIDE_TEMPLATE_DIR"]).expanduser()
    here = Path(__file__).resolve()
    skill = here.parent.parent              # skills/<skill-name>
    for cand in (skill.parent.parent / "assets" / "templates" / "deck",  # リポジトリ直下
                 skill / "assets" / "templates" / "deck"):               # スキル内
        if cand.is_dir():
            return cand
    return skill.parent.parent / "assets" / "templates" / "deck"


def theme_part(prs):
    """スライドマスターに紐づくテーマのパートを返す。"""
    master = prs.slide_masters[0]
    return master.part.part_related_by(RT.THEME)


def theme_root(part):
    """テーマは汎用パートなので _element を持たない。blob を自分で解析する。

    getattr(part, "_element", None) で済ませると、None が返るだけで
    何も起きず、変更したつもりで変わらない。ここは blob を通す。
    """
    return etree.fromstring(part.blob)


def write_theme(part, root):
    part._blob = etree.tostring(root, xml_declaration=True,
                                encoding="UTF-8", standalone=True)


def set_theme_fonts(prs, latin, ea):
    """テーマの見出し用・本文用フォントを差し替える。

    日本語は ea（East Asian）に入る。latin だけ変えても日本語は変わらない。
    """
    part = theme_part(prs)
    root = theme_root(part)
    scheme = root.find(f".//{A}fontScheme")
    for kind in ("majorFont", "minorFont"):
        node = scheme.find(f"{A}{kind}")
        node.find(f"{A}latin").set("typeface", latin)
        ea_node = node.find(f"{A}ea")
        if ea_node is None:
            ea_node = node.makeelement(f"{A}ea", {})
            node.insert(1, ea_node)
        ea_node.set("typeface", ea)
    write_theme(part, root)


def set_theme_colors(prs, colors):
    """テーマの配色を差し替える。キーは dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink。"""
    part = theme_part(prs)
    root = theme_root(part)
    scheme = root.find(f".//{A}clrScheme")
    for child in scheme:
        name = child.tag.split("}")[-1]
        if name not in colors:
            continue
        # sysClr（システム色）が入っていることがあるので srgbClr に置き換える
        for c in list(child):
            child.remove(c)
        srgb = child.makeelement(f"{A}srgbClr", {"val": colors[name]})
        child.append(srgb)
    write_theme(part, root)


def build(outdir: Path, name: str, latin: str, ea: str, colors: dict, note: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)   # 16:9
    prs.slide_height = Inches(7.5)
    set_theme_fonts(prs, latin, ea)
    set_theme_colors(prs, colors)
    prs.core_properties.title = name
    prs.core_properties.comments = note
    path = outdir / f"{name}.pptx"
    prs.save(path)
    print(f"  {path.name}  フォント: {ea} / {latin}")
    return path


# 報告・提案資料向けの抑えた配色。accent1 を主役にする
STANDARD = {
    "dk1": "333333", "lt1": "FFFFFF", "dk2": "1F4E79", "lt2": "F2F5FA",
    "accent1": "1F4E79", "accent2": "4472C4", "accent3": "BDD7EE",
    "accent4": "808080", "accent5": "C00000", "accent6": "D9E2F3",
    "hlink": "1F4E79", "folHlink": "808080",
}

# 印刷・白黒コピー向け。色に頼らず濃淡で差を付ける
MONO = {
    "dk1": "1A1A1A", "lt1": "FFFFFF", "dk2": "3C3C3C", "lt2": "F5F5F5",
    "accent1": "3C3C3C", "accent2": "6E6E6E", "accent3": "A6A6A6",
    "accent4": "D0D0D0", "accent5": "1A1A1A", "accent6": "EDEDED",
    "hlink": "3C3C3C", "folHlink": "6E6E6E",
}


def main():
    outdir = Path(sys.argv[1]) if len(sys.argv) > 1 else template_dir()
    outdir.mkdir(parents=True, exist_ok=True)
    print(f"汎用の資料テンプレートを生成: {outdir}")

    build(outdir, "00_汎用-標準", "Arial", "游ゴシック", STANDARD,
          "報告・提案資料の既定。濃紺を主役にした4色構成")
    build(outdir, "01_汎用-モノクロ", "Arial", "游ゴシック", MONO,
          "白黒印刷・コピー向け。色でなく濃淡で差を付ける")

    print("完了: 2 件")
    print("\n利用者のテンプレートは、この隣に置く。名前は変えなくてよい。")


if __name__ == "__main__":
    main()

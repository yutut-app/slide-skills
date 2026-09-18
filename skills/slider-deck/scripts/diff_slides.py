#!/usr/bin/env python3
"""2つの pptx を比べ、**どのスライドが変わったかを機械的に出す。**

**直した後、変えた枚を確かめるときに使う。**手で直された pptx と HTML の差は `pptx_diff_html.py`。

    python3 scripts/diff_slides.py <前>.pptx <後>.pptx

直したスライドだけが変わっていることの証明に使う。
「3枚目を直した」と報告して実は5枚目も触っていた、を防ぐ。

**手直しのたびに実行し、報告に貼る。** 触っていないことは、
触っていないと言うだけでは証明にならない。

判定は正規化した XML で行う。python-pptx は保存時に XML を書き直すため、
バイト比較では全スライドが「変わった」と出てしまう。
"""

import sys
import zipfile
from pathlib import Path

from lxml import etree


def slide_names(z: zipfile.ZipFile):
    names = [n for n in z.namelist()
             if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    return sorted(names, key=lambda n: int("".join(c for c in Path(n).stem if c.isdigit())))


def canon(blob: bytes) -> bytes:
    return etree.tostring(etree.fromstring(blob), method="c14n2", strip_text=True)


def texts(blob: bytes):
    """スライドの文字を順に返す。差分の中身を説明するために使う。"""
    root = etree.fromstring(blob)
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    return [e.text for e in root.iter(f"{A}t") if e.text and e.text.strip()]


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    before, after = Path(sys.argv[1]), Path(sys.argv[2])
    for p in (before, after):
        if not p.exists():
            sys.exit(f"見つからない: {p}")

    a, b = zipfile.ZipFile(before), zipfile.ZipFile(after)
    na, nb = slide_names(a), slide_names(b)

    print(f"前: {before.name}  {len(na)} 枚")
    print(f"後: {after.name}  {len(nb)} 枚\n")

    changed, same = [], []
    for i in range(max(len(na), len(nb))):
        no = i + 1
        if i >= len(na):
            changed.append((no, "追加", ""))
            continue
        if i >= len(nb):
            changed.append((no, "削除", ""))
            continue
        ba, bb = a.read(na[i]), b.read(nb[i])
        if ba == bb or canon(ba) == canon(bb):
            same.append(no)
            continue
        ta, tb = texts(ba), texts(bb)
        if ta != tb:
            added = [x for x in tb if x not in ta]
            removed = [x for x in ta if x not in tb]
            detail = []
            if removed:
                detail.append("削除: " + " / ".join(x[:18] for x in removed[:3]))
            if added:
                detail.append("追加: " + " / ".join(x[:18] for x in added[:3]))
            changed.append((no, "文字が変わった", "、".join(detail)))
        else:
            changed.append((no, "見た目が変わった", "文字は同じ。位置・色・書式のいずれか"))

    if not changed:
        print("変わったスライドは無い。")
        return 0

    print("| スライド | 変化 | 内容 |")
    print("|---|---|---|")
    for no, kind, detail in changed:
        print(f"| {no} | {kind} | {detail} |")
    print(f"\n変わっていないスライド: "
          f"{', '.join(str(n) for n in same) if same else 'なし'}")
    print("\n**意図した以外のスライドが並んでいたら、そこを調べる。**")
    return 0


if __name__ == "__main__":
    sys.exit(main())

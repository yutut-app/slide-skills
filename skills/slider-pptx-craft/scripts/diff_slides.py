#!/usr/bin/env python3
"""2つの pptx を比べ、**どのスライドが変わったかを機械的に出す。**

**直した後、変えた枚を確かめるときに使う。**手で直された pptx と HTML の差は `pptx_diff_html.py`。

    python3 scripts/diff_slides.py <前>.pptx <後>.pptx
    python3 scripts/diff_slides.py <前>.pptx <後>.pptx --format   # 強調（太字・色・下線）と改行の増減も出す

**`--format` を、人が直した版を土台にするときは必ず付ける。**文字が同じでも、
人が付けた赤太字が消えていることがある。

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


A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def emphasis(blob: bytes):
    """強調された文字（太字・斜体・下線・色）と改行の数。**人が付けた書式を数えるため。**"""
    root = etree.fromstring(blob)
    marks, brs = [], len(list(root.iter(f"{A_NS}br")))
    for r in root.iter(f"{A_NS}r"):
        rpr = r.find(f"{A_NS}rPr")
        t = "".join(x.text or "" for x in r.iter(f"{A_NS}t")).strip()
        if rpr is None or not t:
            continue
        f = []
        if rpr.get("b") == "1": f.append("太字")
        if rpr.get("i") == "1": f.append("斜体")
        if rpr.get("u") not in (None, "none"): f.append("下線")
        c = rpr.find(f"{A_NS}solidFill/{A_NS}srgbClr")
        if c is not None: f.append("色" + c.get("val"))
        if f:
            marks.append((t, "・".join(f)))
    return marks, brs


def format_diff(ba: bytes, bb: bytes):
    ma, ra = emphasis(ba)
    mb, rb = emphasis(bb)
    gone = [m for m in ma if m not in mb]
    new = [m for m in mb if m not in ma]
    out = []
    if gone:
        out.append("**強調が消えた**: " + " / ".join(f"「{t[:14]}」{f}" for t, f in gone[:3]))
    if new:
        out.append("強調が増えた: " + " / ".join(f"「{t[:14]}」{f}" for t, f in new[:3]))
    if ra != rb:
        out.append(f"**改行 {ra}→{rb}**")
    return "、".join(out)


def main():
    argv = [a for a in sys.argv[1:] if a != "--format"]
    want_fmt = "--format" in sys.argv
    if len(argv) != 2:
        sys.exit(__doc__)
    before, after = Path(argv[0]), Path(argv[1])
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
        fd = format_diff(ba, bb) if want_fmt else ""
        if ta != tb:
            added = [x for x in tb if x not in ta]
            removed = [x for x in ta if x not in tb]
            detail = []
            if removed:
                detail.append("削除: " + " / ".join(x[:18] for x in removed[:3]))
            if added:
                detail.append("追加: " + " / ".join(x[:18] for x in added[:3]))
            if fd:
                detail.append(fd)
            changed.append((no, "文字が変わった", "、".join(detail)))
        else:
            changed.append((no, "見た目が変わった", fd or "文字は同じ。位置・色・書式のいずれか"))

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

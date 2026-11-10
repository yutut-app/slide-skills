#!/usr/bin/env python3
"""既存の pptx を、**人が付けた書式（色・太さ・大きさ・改行）を保ったまま**局所的に直す。

    python3 scripts/pptx_edit.py formats <pptx> [--slide N]                  # 直す前に、強調と改行を一覧にする
    python3 scripts/pptx_edit.py text  <pptx> --slide N --find 旧 --replace 新 -o <出力>
    python3 scripts/pptx_edit.py image <pptx> --slide N --pic 1 --file 新.png -o <出力>
    python3 scripts/pptx_edit.py lang  <pptx> -o <出力>                       # 日本語の run の校正言語を ja-JP に
    python3 scripts/pptx_edit.py notes <pptx> --slide N --text-file notes.txt -o <出力>

**同じ物を slider-pptx-craft と slider-script に置く**（`tools/sync_check.py` で一致を見る）。

## なぜ道具にするか

run 単位で手で書き換えると、**文字が重なり、人が付けた強調が消え、枠から溢れる。**
規約（「書式を保つ」）は書いてあっても、手で XML を触ると壊れる。だから境目を機械で守る。

| 置き換える範囲 | 動き |
|---|---|
| 1つの run の中 | その run の中だけ替える。**書式はそのまま** |
| 書式が同じ run をまたぐ | 先頭の run にまとめる（書式は同じなので失われない） |
| **書式が違う run をまたぐ**（赤太字の境目など） | **止める（終了3）。**どこで書式が切り替わるかを出す。範囲を分けて替える |
| **改行をまたぐ** | **止める（終了3）。**改行は人が入れたもの。消さない |

書き出しの前後:
- `-o` の先に既にファイルがあれば、`--force` が無いと書かない。`--force` でも **`guard_edit.py check` を通す**
- 書き出したら **`guard_edit.py record` で控える**（次に上書きするとき、比べる元になる）
- 書き換えた run に日本語があれば `lang="ja-JP"` を付ける（**赤い波線を出さない**）
"""

import argparse
import copy
import re
import subprocess
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Emu  # noqa: F401

sys.path.insert(0, str(Path(__file__).resolve().parent))
import checked  # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
JA = re.compile(r"[぀-ヿ㐀-鿿＀-￯]")
HERE = Path(__file__).resolve().parent


# ---------------------------------------------------------------- 走査
def rtext(el):
    """要素の中の文字。**itertext() は使わない**（python-pptx の要素では同じ文字が2回返る）。"""
    return "".join(t.text or "" for t in el.iter(f"{A}t"))


def iter_text_frames(shapes):
    for sh in shapes:
        if sh.shape_type == 6 and hasattr(sh, "shapes"):       # グループ
            yield from iter_text_frames(sh.shapes)
            continue
        if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
            yield sh.name, sh.text_frame
        if getattr(sh, "has_table", False) and sh.has_table:
            for r, row in enumerate(sh.table.rows):
                for c, cell in enumerate(row.cells):
                    yield f"{sh.name}[{r + 1},{c + 1}]", cell.text_frame


def fmt_of(r):
    """run の書式を、人が読める短い形にする。比べるときは rPr の XML そのものを使う。"""
    rpr = r.find(f"{A}rPr")
    if rpr is None:
        return ""
    parts = []
    if rpr.get("b") == "1":
        parts.append("太字")
    if rpr.get("i") == "1":
        parts.append("斜体")
    if rpr.get("u") not in (None, "none"):
        parts.append("下線")
    if rpr.get("sz"):
        parts.append(f"{int(rpr.get('sz')) / 100:g}pt")
    clr = rpr.find(f"{A}solidFill/{A}srgbClr")
    if clr is not None:
        parts.append("色" + clr.get("val"))
    sch = rpr.find(f"{A}solidFill/{A}schemeClr")
    if sch is not None:
        parts.append("色" + sch.get("val"))
    return "・".join(parts)


def rpr_key(r):
    rpr = r.find(f"{A}rPr")
    if rpr is None:
        return b""
    k = copy.deepcopy(rpr)
    for att in ("lang", "altLang", "dirty", "err", "noProof", "smtClean"):   # 書式でないもの
        k.attrib.pop(att, None)
    return etree.tostring(k, method="c14n")


def set_lang(r, lang="ja-JP"):
    rpr = r.find(f"{A}rPr")
    if rpr is None:
        rpr = etree.SubElement(r, f"{A}rPr")
        r.remove(rpr)
        r.insert(0, rpr)
    rpr.set("lang", lang)
    rpr.attrib.pop("err", None)


# ---------------------------------------------------------------- 書き出し
def guard_and_save(prs, out: Path, force: bool):
    if out.exists():
        if not force:
            sys.exit(f"既にある: {out}\n**書かない。**人が直したものかもしれない。"
                     f"確かめる: python3 {HERE / 'guard_edit.py'} check {out}  → 変わっていなければ --force")
        r = subprocess.run([sys.executable, str(HERE / "guard_edit.py"), "check", str(out)])
        if r.returncode != 0:
            sys.exit("**guard_edit が止めた。上書きしない。**変わった枚を利用者に見せて、どうするか決めてもらう")
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    subprocess.run([sys.executable, str(HERE / "guard_edit.py"), "record", str(out)],
                   stdout=subprocess.DEVNULL)
    print(f"書き出した: {out}（控えた。次に上書きするときの比べる元）")


# ---------------------------------------------------------------- formats
def cmd_formats(prs, args):
    print("| 枚 | 枠 | 文字 | 書式 |")
    print("|---|---|---|---|")
    n = 0
    for no, slide in enumerate(prs.slides, 1):
        if args.slide and no != args.slide:
            continue
        for name, tf in iter_text_frames(slide.shapes):
            for p in tf.paragraphs:
                brs = len(p._p.findall(f"{A}br"))
                for r in p._p.findall(f"{A}r"):
                    f = fmt_of(r)
                    t = rtext(r)
                    if t.strip() and any(k in f for k in ("太字", "色", "下線", "斜体")):
                        print(f"| {no} | {name} | {t[:24]} | {f} |"); n += 1
                if brs:
                    txt = rtext(p._p)
                    print(f"| {no} | {name} | {txt[:24]}… | 改行 {brs} 箇所（人が入れたもの。消さない） |"); n += 1
    return checked.summary("pptx_edit formats", len(prs.slides._sldIdLst), "枚", 0,
                           {"強調・改行": f"{n} 箇所"})


# ---------------------------------------------------------------- text
def cmd_text(prs, args):
    slide = prs.slides[args.slide - 1]
    hits, broken = [], []
    for name, tf in iter_text_frames(slide.shapes):
        for p in tf.paragraphs:
            kids = [k for k in p._p if k.tag in (f"{A}r", f"{A}br", f"{A}fld")]
            segs, pos = [], 0
            for k in kids:
                t = "\n" if k.tag == f"{A}br" else rtext(k)
                segs.append((k, pos, pos + len(t)))
                pos += len(t)
            full = "".join("\n" if k.tag == f"{A}br" else rtext(k) for k in kids)
            if args.find not in full and args.find in full.replace("\n", ""):
                broken.append(name)
            start = 0
            while (i := full.find(args.find, start)) >= 0:
                hits.append((name, segs, i, i + len(args.find)))
                start = i + 1
    if not hits and broken:
        print(f"**止めた（改行をまたぐ）**: {', '.join(broken)}。改行は人が入れたもの。改行の前後で2回に分けて替える")
        return 3
    if not hits:
        sys.exit(f"{args.slide}枚目に「{args.find}」が無い。**表記（全角・半角・改行）を `formats` で確かめる**")
    if len(hits) > 1 and not args.nth:
        sys.exit(f"{args.slide}枚目に「{args.find}」が {len(hits)} 箇所ある（"
                 + "、".join(h[0] for h in hits) + "）。`--nth` で何番目かを指定する")
    name, segs, s, e = hits[(args.nth or 1) - 1]
    over = [(k, a, b) for k, a, b in segs if a < e and b > s]
    if any(k.tag == f"{A}br" for k, _, _ in over):
        print(f"**止めた（改行をまたぐ）**: {name}。改行は人が入れたもの。改行の前後で2回に分けて替える")
        return 3
    keys = {rpr_key(k) for k, _, _ in over}
    if len(keys) > 1:
        print(f"**止めた（書式の境目をまたぐ）**: {name}。範囲の中で書式が切り替わる:")
        for k, a, b in over:
            print(f"- 「{rtext(k)}」 {fmt_of(k) or '既定'}")
        print("**範囲を書式ごとに分けて替える。**どちらの書式にするかは、こちらで決めない")
        return 3
    first, fa, _ = over[0]
    last, la, _ = over[-1]
    t_first = rtext(first)
    t_last = rtext(last)
    new = t_first[: s - fa] + args.replace + t_last[e - la:]
    if fmt_of(first) and len(args.replace) > len(args.find) and args.find in args.replace:
        print(f"**注意（書式が広がる）**: 足した文字「{args.replace.replace(args.find, '', 1)}」も {fmt_of(first)} になる。"
              "広げてよいかは修正案の表の「書式への影響」で確かめる。広げないなら、強調の外の文字を `--find` に選んで足す")
    first.find(f"{A}t").text = new
    for k, _, _ in over[1:]:
        k.getparent().remove(k)
    if JA.search(new):
        set_lang(first)
    print(f"替えた: {args.slide}枚目 {name}「{args.find}」→「{args.replace}」（書式: {fmt_of(first) or '既定'}、そのまま）")
    guard_and_save(prs, Path(args.out), args.force)
    return 0


# ---------------------------------------------------------------- image
def cmd_image(prs, args):
    from PIL import Image
    slide = prs.slides[args.slide - 1]
    pics = [sh for sh in slide.shapes if sh.shape_type == 13]
    if not pics or args.pic > len(pics):
        sys.exit(f"{args.slide}枚目の画像は {len(pics)} 枚。--pic は 1〜{len(pics)}")
    pic = pics[args.pic - 1]
    _, rid = pic.part.get_or_add_image_part(args.file)
    pic._element.blipFill.blip.rEmbed = rid               # 枠はそのまま、画像だけ替わる
    w, h = Image.open(args.file).size
    frame, img = pic.width / pic.height, w / h
    pic.crop_left = pic.crop_right = pic.crop_top = pic.crop_bottom = 0
    if img > frame:
        pic.crop_left = pic.crop_right = (1 - frame / img) / 2
        cut = "左右"
    else:
        pic.crop_top = pic.crop_bottom = (1 - img / frame) / 2
        cut = "上下"
    print(f"替えた: {args.slide}枚目 {pic.name}。枠はそのまま、比が違う分は{cut}を中央で切った。"
          "**端の表・凡例が切れていないかを絵で見る**")
    guard_and_save(prs, Path(args.out), args.force)
    return 0


# ---------------------------------------------------------------- lang
def cmd_lang(prs, args):
    n = 0
    for slide in prs.slides:
        for _, tf in iter_text_frames(slide.shapes):
            for r in tf._txBody.iter(f"{A}r"):
                rpr = r.find(f"{A}rPr")
                if JA.search(rtext(r)) and (rpr is None or rpr.get("lang") != args.lang):
                    set_lang(r, args.lang); n += 1
    print(f"校正の言語を {args.lang} にした run: {n} 個")
    guard_and_save(prs, Path(args.out), args.force)
    return 0


# ---------------------------------------------------------------- notes
def cmd_notes(prs, args):
    slide = prs.slides[args.slide - 1]
    lines = [l.rstrip() for l in Path(args.text_file).read_text(encoding="utf-8").splitlines() if l.strip()]
    tf = slide.notes_slide.notes_text_frame
    tf.text = lines[0] if lines else ""
    for l in lines[1:]:
        tf.add_paragraph().text = l
    for r in tf._txBody.iter(f"{A}r"):
        if JA.search(rtext(r)):
            set_lang(r)
    print(f"ノートを書いた: {args.slide}枚目 {len(lines)} 行")
    guard_and_save(prs, Path(args.out), args.force)
    return 0


def main():
    ap = argparse.ArgumentParser(description="既存の pptx を、書式を保って局所的に直す")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("formats", "text", "image", "lang", "notes"):
        s = sub.add_parser(name)
        s.add_argument("pptx")
        s.add_argument("--slide", type=int)
        if name != "formats":
            s.add_argument("-o", "--out", required=True)
            s.add_argument("--force", action="store_true")
    sp = sub.choices
    sp["text"].add_argument("--find", required=True)
    sp["text"].add_argument("--replace", required=True)
    sp["text"].add_argument("--nth", type=int)
    sp["image"].add_argument("--pic", type=int, default=1)
    sp["image"].add_argument("--file", required=True)
    sp["lang"].add_argument("--lang", default="ja-JP")
    sp["notes"].add_argument("--text-file", required=True)
    args = ap.parse_args()
    if args.cmd in ("text", "image", "notes") and not args.slide:
        ap.error("--slide が要る")
    prs = Presentation(args.pptx)
    return {"formats": cmd_formats, "text": cmd_text, "image": cmd_image,
            "lang": cmd_lang, "notes": cmd_notes}[args.cmd](prs, args)


if __name__ == "__main__":
    sys.exit(main())

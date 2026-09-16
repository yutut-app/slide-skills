#!/usr/bin/env python3
"""資料テンプレート（.pptx）から、中の画像を取り出す。

    python3 scripts/extract_assets.py <テンプレ>.pptx --list
    python3 scripts/extract_assets.py <テンプレ>.pptx -o <テンプレ>/images

## 何のためにあるか

**HTML テンプレートの見た目は、100% HTML 側で決まる。**
変換は python-pptx の空ファイルを土台にするので、
**元の .pptx のマスターもテーマもロゴも、出力には一切入らない**
（`<slider-deck>/scripts/html2pptx.py` の `build_absolute`）。

ロゴが出るのは、HTML が `<img src="images/logo.png">` を置いているからで、
**その画像が仮物なら、出てくる資料も仮物のまま。**
実際に、実物のテンプレートには入っているロゴが、
HTML テンプレート側では 1.9KB のプレースホルダのままだった。

**実物に合わせるとは、実物の画像を `images/` に入れること。**
手で PowerPoint を開いて画像を取り出すと、
**どれがどこで使われている画像かが分からなくなる。**ここで機械的に出す。

## 出すもの

画像ごとに、**大きさと、どこで使われているか**を出す。

| 使われ場所 | 意味 |
|---|---|
| `master` | スライドマスター。**全ページに出る。**ロゴはたいていここ |
| `layout:<n>` | レイアウト。その種別のページに出る |
| `slide:<n>` | そのページだけ |

**マスターかレイアウトで使われている大きな画像が、ふつうロゴ。**
ただし**決め打ちしない。**大きさと使われ場所を見て、人が選ぶ。

## 入れ替えるとき

**上書きする前に、今あるものを見る。** 名前が同じでも中身が違う。

    python3 scripts/extract_assets.py <テンプレ>.pptx -o /tmp/assets
    # 出た表を見て、ロゴを選ぶ
    cp /tmp/assets/<選んだもの>.png <テンプレ>/images/logo.png

**入れ替えたら指紋を取り直す**（`handoff.py --fingerprint`）。
指紋は `images/` も見ているので、値が変わる。
**両方の手元で同じ値になるまで、渡す側と受け取る側で突き合わせる。**

外に出さない。ローカルにファイルを書くだけで、公開もアップロードもしない。
"""

import argparse
import posixpath
import re
import shutil
import struct
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked

MEDIA_RE = re.compile(r"^ppt/media/")
EMBED_RE = re.compile(r'r:(?:embed|link|pict|id)="([^"]+)"')
REL_RE = re.compile(r'Id="([^"]+)"[^>]*Target="([^"]+)"')


def part_kind(name: str):
    """XML パートの名前から、使われ場所の呼び名を作る。"""
    m = re.match(r"ppt/slideMasters/slideMaster(\d+)\.xml$", name)
    if m:
        return "master" if m.group(1) == "1" else f"master:{m.group(1)}"
    m = re.match(r"ppt/slideLayouts/slideLayout(\d+)\.xml$", name)
    if m:
        return f"layout:{m.group(1)}"
    m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
    if m:
        return f"slide:{m.group(1)}"
    if name == "ppt/theme/theme1.xml":
        return "theme"
    return None


def usage_map(z: zipfile.ZipFile):
    """画像 → それを使っているパートの並び。

    **関係は .rels にしか無い。**XML 本体には rId しか書かれていないので、
    片方だけ読んでも、どの画像がどこで使われているかは出ない。
    """
    used = {}
    for name in z.namelist():
        kind = part_kind(name)
        if not kind:
            continue
        try:
            body = z.read(name).decode("utf-8", "replace")
        except KeyError:
            continue
        rids = set(EMBED_RE.findall(body))
        if not rids:
            continue
        rels_name = posixpath.join(posixpath.dirname(name), "_rels",
                                   posixpath.basename(name) + ".rels")
        if rels_name not in z.namelist():
            continue
        rels = z.read(rels_name).decode("utf-8", "replace")
        for rid, target in REL_RE.findall(rels):
            if rid not in rids or "media/" not in target:
                continue
            media = posixpath.normpath(
                posixpath.join(posixpath.dirname(name), target))
            used.setdefault(media, []).append(kind)
    return used


def png_size(b: bytes):
    if b[:8] == b"\x89PNG\r\n\x1a\n" and b[12:16] == b"IHDR":
        return struct.unpack(">II", b[16:24])
    return None


def jpeg_size(b: bytes):
    if b[:2] != b"\xff\xd8":
        return None
    i = 2
    while i + 9 < len(b):
        if b[i] != 0xFF:
            i += 1
            continue
        marker, seg = b[i + 1], struct.unpack(">H", b[i + 2:i + 4])[0]
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            h, w = struct.unpack(">HH", b[i + 5:i + 9])
            return w, h
        i += 2 + seg
    return None


def dimensions(b: bytes):
    """**分からないものは分からないと出す。**推測で埋めない。"""
    return png_size(b) or jpeg_size(b)


def main():
    ap = argparse.ArgumentParser(
        description="テンプレートの .pptx から画像を取り出す（ローカルのみ。公開しない）")
    ap.add_argument("pptx", help="テンプレートの .pptx / .potx")
    ap.add_argument("-o", "--out", help="取り出し先ディレクトリ")
    ap.add_argument("--list", action="store_true", help="一覧を出すだけ。書かない")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        sys.exit(f"見つからない: {src}")
    if not args.list and not args.out:
        ap.error("-o で取り出し先を渡す（見るだけなら --list）")

    try:
        z = zipfile.ZipFile(src)
    except zipfile.BadZipFile:
        sys.exit(f"pptx として開けない: {src}\n"
                 "  **.pptx / .potx か確かめる。**中身が壊れていることもある")

    media = sorted(n for n in z.namelist() if MEDIA_RE.match(n))
    used = usage_map(z)

    outdir = Path(args.out) if args.out else None
    if outdir and not args.list:
        outdir.mkdir(parents=True, exist_ok=True)

    rows, unknown = [], 0
    for name in media:
        b = z.read(name)
        wh = dimensions(b)
        size = f"{wh[0]}×{wh[1]}px" if wh else "**不明**"
        if not wh:
            unknown += 1
        where = used.get(name, [])
        base = posixpath.basename(name)
        if outdir and not args.list:
            with z.open(name) as fsrc, open(outdir / base, "wb") as fdst:
                shutil.copyfileobj(fsrc, fdst)
        rows.append((base, size, f"{len(b)/1024:.1f}KB",
                     ", ".join(where) if where else "**どこにも無い**"))

    where = "一覧のみ（書いていない）" if args.list else f"取り出し先 {outdir}/"
    print(f"画像: {len(media)} 件  {where}")
    if not media:
        print("\n**画像が1つも入っていない。**ロゴが図形で描かれていることがある。")
        print("その場合はここでは取り出せない。**PowerPoint 上で見て確かめる。**")
    else:
        print()
        print("| ファイル | 大きさ | 容量 | 使われ場所 |")
        print("|---|---|---|---|")
        for r in rows:
            print("| " + " | ".join(r) + " |")
        print()
        print("**master / layout で使われている大きなものが、ふつうロゴ。**")
        print("**決め打ちしない。**大きさと使われ場所を見て選ぶ。")
        if not args.list:
            print(f"\n  cp {outdir}/<選んだもの> <テンプレ>/images/logo.png")
            print("**入れ替えたら指紋を取り直す**（handoff.py --fingerprint）。"
                  "渡す側と受け取る側で突き合わせる。")

    return checked.summary("extract_assets", len(media), "件", unknown, {
        "元": src.name,
        "書き込み": "なし（一覧のみ）" if args.list else str(outdir),
    })


if __name__ == "__main__":
    sys.exit(main())

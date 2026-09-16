#!/usr/bin/env python3
"""HTML 一式を、別の資料テンプレートに載せ替える。

    python3 scripts/reskin_html.py --list
    python3 scripts/reskin_html.py work/ --to <テンプレ名> -o work_<テンプレ名> --dry-run
    python3 scripts/reskin_html.py work/ --to <テンプレ名> -o work_<テンプレ名>

**載せ替えるのは HTML。pptx ではない。**
pptx は派生物であり、手で直しても次の変換で消える（slider-craft `48_html.md`）。
**正本である HTML を載せ替え、pptx は変換し直す。**

## 何を入れ替えているか

資料テンプレートの HTML は、**`<style>` が器、`<body>` が中身**に分かれている。
実測: 同じ器の 16:9 版と A4 版は、`<body>` が一致するか、
差はコメントとページ番号だけだった。**クラス名の集合もほぼ一致する。**

    新しい HTML = 移行先テンプレの <head>（= <style>）+ 元の <body>

座標・文字サイズ・スライド寸法はすべて `<style>` 側にあるので、
これだけで寸法も配色も書体も移行先のものになる。

## どのレイアウトに載せるかの決め方

上から順に、決まった時点で止める。

1. `--layout` で指定されていればそれ（**人が打った指定が最優先**）
2. `<meta name="slide-layout" content="003" />` があればそれ（前回の載せ替えが書いた印）
2-2. **ファイル名がレイアウト名と同じならそれ**（`003.html` → `003`）
3. **本文が使っているクラスのうち、そのレイアウトにしか無いものを数えて採点**する
   （全レイアウト共通のヘッダー・フッターのクラスは除く。
   入れたままだと差が消え、ほとんどが「僅差」になる）

3 で決めたときは**採点結果を必ず出す。** 黙って選ぶと、
違うレイアウトに載ったことに気づけない。

**1位と2位の差が1点以下なら「要確認」として指摘に数える。**
中身ページの派生（目次・リードあり・フリー）はクラスをほとんど共有するので、
採点だけでは分けきれない。実測で 4/5 枚が1点差だった。
**当たっていても、外れたときに気づけない差。** 要確認と出た枚は、
絵を見て確かめるか `--layout` で指定する。

## 載せ替えでは直らないもの

| | なぜ |
|---|---|
| **溢れ** | スライド寸法と文字サイズが変わる。**載せ替え後に必ず検査する** |
| 本文に直書きしたページ番号 | `<body>` 側にあるので、そのまま移る |
| 移行先に無い図版 | `images/` の中身が違えば出ない。**欠落として報告する** |

外に出さない。ローカルにファイルを書くだけで、公開もアップロードもしない。
"""

import argparse
import os
import re
import shutil
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked

SKILL_ROOT = Path(__file__).resolve().parent.parent
MARKER = 'slide-layout'


# ---------------------------------------------------------------------------
# テンプレートの探索
# ---------------------------------------------------------------------------

def template_dir() -> Path:
    """資料テンプレートの置き場所。**git に入っていない**（20_templates.md）。

    探索の順は `20_templates.md` の表と同じにする。**規約と実装を2箇所で決めない。**
    実際に、規約は `SLIDE_TEMPLATE_DIR` を1番目に挙げているのに、
    ここが読んでいなかった（**別の機械に配ると、規約どおりにしても効かない**）。
    """
    if os.environ.get("SLIDE_TEMPLATE_DIR"):
        return Path(os.environ["SLIDE_TEMPLATE_DIR"]).expanduser()
    for base in (SKILL_ROOT.parent.parent, SKILL_ROOT.parent):
        d = base / "assets" / "templates" / "deck"
        if d.is_dir():
            return d
    sys.exit("テンプレートの置き場所が見つからない: assets/templates/deck/\n"
             "  クローン直後には無い（テンプレートは git に入れていない）。\n"
             "  渡されたテンプレートのフォルダを、そこに作って置く。\n"
             "  別の場所に置くなら環境変数 SLIDE_TEMPLATE_DIR を向ける。")


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def find_templates() -> list[Path]:
    """HTML テンプレート = template-manifest.md を持つディレクトリ。"""
    return sorted((d for d in template_dir().iterdir()
                   if d.is_dir() and (d / "template-manifest.md").exists()),
                  key=lambda p: nfc(p.name))


def resolve_template(name: str) -> Path:
    cands = find_templates()
    for d in cands:
        if nfc(d.name) == nfc(name):
            return d
    hit = [d for d in cands if nfc(name) in nfc(d.name)]
    if len(hit) == 1:
        return hit[0]
    if not hit:
        sys.exit(f"テンプレートが無い: {name}\n" + describe_templates())
    sys.exit("どれか決まらない: " + " / ".join(d.name for d in hit))


def slide_size(text: str) -> str:
    m = re.search(r"\.slide\s*\{[^}]*?width:\s*([\d.]+)px[^}]*?height:\s*([\d.]+)px",
                  text, re.S)
    return f"{m.group(1)}x{m.group(2)}px" if m else "不明"


def layouts_of(tpl: Path) -> list[Path]:
    """載せ先の候補。print.html はビューアなので除く。"""
    return sorted((p for p in tpl.glob("*.html") if p.name != "print.html"),
                  key=lambda p: nfc(p.name))


def describe_templates() -> str:
    out = ["選べるテンプレート", "", "| テンプレート | レイアウト | スライド寸法 |", "|---|---|---|"]
    for d in find_templates():
        ls = layouts_of(d)
        size = slide_size(ls[0].read_text(encoding="utf-8")) if ls else "不明"
        out.append(f"| {d.name} | {len(ls)}種 | {size} |")
    out.append("")
    out.append("**利用者に選ばせる。こちらで決めない。**")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# HTML の読み書き
# ---------------------------------------------------------------------------

BODY_RE = re.compile(r"(<body[^>]*>)(.*)(</body>)", re.S | re.I)
CSS_COMMENT_RE = re.compile(r"/\*.*?\*/", re.S)
STYLE_RE = re.compile(r"<style[^>]*>(.*?)</style>", re.S | re.I)
HEAD_RE = re.compile(r"(<head[^>]*>)(.*)(</head>)", re.S | re.I)


def body_of(text: str) -> str | None:
    m = BODY_RE.search(text)
    return m.group(2) if m else None


def classes_used(body: str) -> set[str]:
    used = set()
    for m in re.finditer(r'class\s*=\s*"([^"]*)"', body):
        used |= set(m.group(1).split())
    for m in re.finditer(r"class\s*=\s*'([^']*)'", body):
        used |= set(m.group(1).split())
    return used


def classes_defined(text: str) -> set[str]:
    """`<style>` の中で定義されているクラス。**コメントは先に落とす。**"""
    css = " ".join(STYLE_RE.findall(text))
    css = CSS_COMMENT_RE.sub(" ", css)
    css = re.sub(r"\{[^{}]*\}", " { } ", css)      # 宣言部を落とし、セレクタだけ残す
    return set(re.findall(r"\.([A-Za-z][\w-]*)", css))


def images_used(body: str) -> set[str]:
    return set(re.findall(r'src\s*=\s*"([^"]+\.(?:png|jpe?g|gif|svg|webp))"', body, re.I))


def marker_of(text: str) -> str | None:
    m = re.search(rf'<meta\s+name="{MARKER}"\s+content="([^"]+)"', text, re.I)
    return m.group(1) if m else None


def put_marker(text: str, layout: str) -> str:
    """載せたレイアウトを印として残す。次の載せ替えが採点なしで決められる。"""
    tag = f'<meta name="{MARKER}" content="{layout}" />'
    if re.search(rf'<meta\s+name="{MARKER}"[^>]*>', text, re.I):
        return re.sub(rf'<meta\s+name="{MARKER}"[^>]*>', tag, text, count=1, flags=re.I)
    m = HEAD_RE.search(text)
    if not m:
        return text
    return text[:m.start(2)] + f"\n    {tag}" + text[m.start(2):]


# ---------------------------------------------------------------------------
# レイアウトの選定
# ---------------------------------------------------------------------------

def pick_layout(src_text: str, cands: list[Path], forced: str | None,
                stem: str | None = None):
    """(選んだレイアウト, 決め方, 採点表) を返す。"""
    if forced:
        for c in cands:
            if c.stem == forced or c.name == forced:
                return c, f"指定（--layout {forced}）", None
        sys.exit(f"指定されたレイアウトが無い: {forced}\n"
                 "  候補: " + " / ".join(c.stem for c in cands))
    mark = marker_of(src_text)
    if mark:
        for c in cands:
            if c.stem == mark:
                return c, f"印（meta {MARKER}={mark}）", None
    if stem:
        for c in cands:
            if nfc(c.stem) == nfc(stem):
                return c, f"同名（{stem}）", None

    used = classes_used(body_of(src_text) or "")
    defs = {c: classes_defined(c.read_text(encoding="utf-8")) for c in cands}
    # **全レイアウトが持つクラスは採点から除く。**
    # ヘッダー・フッター・共通の枠は、どのレイアウトにもある。
    # 入れたままだと点数が底上げされ、1位と2位の差が消えて
    # ほとんどの枚が「僅差」になる（実測: 5枚中4枚）。**区別できるものだけで数える。**
    common = set.intersection(*defs.values()) if defs else set()
    table = [(len((used & d) - common), c) for c, d in defs.items()]
    table.sort(key=lambda t: (-t[0], nfc(t[1].name)))
    best = table[0][0]
    tied = [c for s, c in table if s == best]
    if best == 0:
        return None, "決まらない（一致するクラスが1つも無い）", table
    if len(tied) > 1:
        return None, "決まらない（同点: " + " / ".join(c.stem for c in tied) + "）", table
    second = table[1][0] if len(table) > 1 else 0
    if best - second < 2:
        # **1点差は確信ではない。** 中身ページの派生（目次・リードあり・フリー）は
        # クラスをほとんど共有するので、採点だけでは分けきれない。
        # 決めはするが**要確認として指摘に数える。** 黙って通さない。
        return table[0][1], f"採点（{best}対{second}／**要確認**）", table
    return table[0][1], f"採点（{best}クラス一致）", table


# ---------------------------------------------------------------------------
# 本体
# ---------------------------------------------------------------------------

def expand(paths: list[str]) -> list[Path]:
    """ディレクトリと print.html を、スライドの HTML の並びに展開する。"""
    out: list[Path] = []
    for s in paths:
        p = Path(s)
        if p.is_dir():
            out += [q for q in sorted(p.glob("*.html"), key=lambda x: nfc(x.name))
                    if q.name != "print.html"]
            continue
        if not p.exists():
            sys.exit(f"見つからない: {p}")
        if p.name == "print.html":
            frames = re.findall(r'<iframe[^>]*src="([^"]+)"', p.read_text(encoding="utf-8"))
            out += [p.parent / f for f in frames]
        else:
            out.append(p)
    if not out:
        sys.exit("載せ替える HTML が1つも無い")
    return out


def write_print_html(tpl: Path, outdir: Path, names: list[str]):
    """移行先のビューアを雛形にして、並びだけ差し替える。"""
    src = tpl / "print.html"
    if not src.exists():
        return
    text = src.read_text(encoding="utf-8")
    frames = "\n".join(
        f'    <div class="slide-frame"><iframe src="{n}"></iframe></div>' for n in names)
    text = BODY_RE.sub(lambda m: m.group(1) + "\n" + frames + "\n" + m.group(3), text, count=1)
    (outdir / "print.html").write_text(text, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(
        description="HTML 一式を別の資料テンプレートに載せ替える（ローカルのみ。公開しない）")
    ap.add_argument("html", nargs="*",
                    help="載せ替える HTML。ディレクトリか print.html を渡すと展開する")
    ap.add_argument("--to", help="移行先のテンプレート名")
    ap.add_argument("--layout", help="全ファイルを、このレイアウトに載せる（001 など）")
    ap.add_argument("-o", "--out", help="出力先ディレクトリ")
    ap.add_argument("--list", action="store_true", help="選べるテンプレートを出す")
    ap.add_argument("--dry-run", action="store_true", help="書かずに、割付と欠落だけ出す")
    args = ap.parse_args()

    if args.list:
        print(describe_templates())
        return 0
    if not args.html or not args.to:
        ap.error("HTML と --to を渡す（一覧だけ見るなら --list）")
    if not args.dry_run and not args.out:
        ap.error("-o で出力先を渡す（元を上書きしない）")

    tpl = resolve_template(args.to)
    cands = layouts_of(tpl)
    if not cands:
        sys.exit(f"移行先にレイアウトが無い: {tpl.name}")
    files = expand(args.html)
    outdir = Path(args.out) if args.out else None

    tpl_size = slide_size(cands[0].read_text(encoding="utf-8"))
    rows, missing_cls, missing_img, undecided, unsure = [], {}, {}, 0, 0

    if outdir and not args.dry_run:
        outdir.mkdir(parents=True, exist_ok=True)

    written = []
    for f in files:
        src = f.read_text(encoding="utf-8")
        sbody = body_of(src)
        if sbody is None:
            rows.append((f.name, "**<body> が無い**", "-", "-"))
            undecided += 1
            continue

        lay, how, table = pick_layout(src, cands, args.layout, f.stem)
        if lay is None:
            undecided += 1
            rows.append((f.name, f"**{how}**", "-", "-"))
            if table:
                print(f"  {f.name} の採点: " +
                      ", ".join(f"{c.stem}={s}" for s, c in table), file=sys.stderr)
            continue

        dst = lay.read_text(encoding="utf-8")
        used = classes_used(sbody)
        # **欠落 = 元では効いていたのに、移行先で定義が無いもの。**
        # どちらにも無いもの（外部CSS由来）は載せ替えで変わらないので数えない。
        lost = sorted((used & classes_defined(src)) - classes_defined(dst))
        if lost:
            missing_cls[f.name] = lost

        imgs = sorted(i for i in images_used(sbody)
                      if not (tpl / i).exists())
        if imgs:
            missing_img[f.name] = imgs

        new = BODY_RE.sub(lambda m: m.group(1) + sbody + m.group(3), dst, count=1)
        new = put_marker(new, lay.stem)
        if outdir and not args.dry_run:
            (outdir / f.name).write_text(new, encoding="utf-8")
            written.append(f.name)
        if "要確認" in how:
            unsure += 1
        rows.append((f.name, lay.stem, how, f"欠落 {len(lost)}" if lost else "—"))

    if outdir and not args.dry_run:
        src_img = tpl / "images"
        if src_img.is_dir():
            shutil.copytree(src_img, outdir / "images", dirs_exist_ok=True)
        write_print_html(tpl, outdir, written)

    mode = "下見（書いていない）" if args.dry_run else f"出力 → {outdir}/"
    print(f"載せ替え: {len(files)} 枚 → {tpl.name}（{tpl_size}）  {mode}")
    print()
    print("| HTML | 載せたレイアウト | 決め方 | |")
    print("|---|---|---|---|")
    for r in rows:
        print("| " + " | ".join(str(x) for x in r) + " |")

    findings = undecided + unsure + len(missing_cls) + len(missing_img)

    if unsure:
        print(f"\n**割付が僅差で決まった枚が {unsure} 件ある。** 当たっていても、")
        print("外れたときに気づけない差。**その枚は絵で確かめる**か `--layout` で指定する。")
    if undecided:
        print(f"\n**割付が決まらない枚が {undecided} 件ある。** `--layout` で指定するか、")
        print("移行先にそのレイアウトが無いことを疑う。**そのまま進めない。**")
    if missing_cls:
        print("\n**移行先に定義が無いクラス（元では効いていた）。見た目が崩れる。**")
        for n, cs in missing_cls.items():
            print(f"- {n}: " + ", ".join(cs))
    if missing_img:
        print("\n**移行先に無い画像。出ない。**")
        for n, ims in missing_img.items():
            print(f"- {n}: " + ", ".join(ims))

    print("\n**寸法と文字サイズが変わっている。溢れは載せ替えでは直らない。**")
    print("  必ず次を回して、1枚ずつ自分で見る。")
    print("    python3 <slider-craft>/scripts/html_png.py <出力>/print.html -o qa_html")
    print("    python3 <slider-craft>/scripts/html2pptx.py <出力>/print.html -o deck/deck.pptx")
    print("    python3 <slider-craft>/scripts/fit_check.py deck/deck.pptx")

    return checked.summary("reskin_html", len(rows), "枚", findings, {
        "移行先": tpl.name,
        "寸法": tpl_size,
        "レイアウト": len(cands),
        "書き込み": "なし（下見）" if args.dry_run else str(outdir),
    })


if __name__ == "__main__":
    sys.exit(main())

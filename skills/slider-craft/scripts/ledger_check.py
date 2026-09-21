#!/usr/bin/env python3
"""直した後に、台帳で照合する。**古い語・古い値が0件になったことを機械で確かめる。**

    python3 scripts/ledger_check.py work/*.html --terms terms.md --numbers numbers.md
    python3 scripts/ledger_check.py work/*.html --history      # 修正の経緯が漏れた印

## なぜ要るか

**「直しました」は、古いものが残っていないことの証明にならない。**
置換は HTML の空白（`&nbsp;`）や改行で当たらないことがあり、
数値は同じ数字が別の項目にも出る。**目で探すと必ず取り残す。**

| 照合するもの | 出どころ | 残っていたら |
|---|---|---|
| 使わない語 | 用語台帳 `terms.md` の「使わない語」列 | **異常終了**（終了コード 1） |
| 以前の値 | 数値台帳 `numbers.md` の「以前の値」列 | **異常終了** |
| 修正の経緯が漏れた印 | 「〜ではなく」「まだ」「ここでは」「前回」など | 見直し候補として出す（落とさない） |

本文は**見える文字だけ**を読む（タグ・スタイル・スクリプトは除く。`&nbsp;` は空白にする）。
数値は**前後が数字でないものだけ**を拾う（「62%」を探して「162%」に当てない）。
"""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
import lxml.html as LH

HISTORY = ["ではなく", "ではない", "まだ", "ここでは", "前回", "改めて", "再度"]


def table_rows(md: Path):
    """Markdown の表を、見出しの名前をキーにした辞書の並びで返す。"""
    rows, head = [], None
    for line in md.read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            head = None if not rows and head is None else head
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if set("".join(cells)) <= set("-: "):
            continue                               # 区切りの行
        if head is None:
            head = cells
            continue
        rows.append(dict(zip(head, cells)))
    return rows


def split_words(cell: str):
    return [w.strip() for w in re.split(r"[,、，]", cell or "") if w.strip()]


def slides_text(html: Path):
    """[(枚番号, 見える文字)]。1ファイル1枚なら枚番号は 1。"""
    # **UTF-8 で明示的に読む。**`<meta charset>` の無い HTML を lxml は latin-1 で読み、
    # 文字化けした本文と照合して**黙って0件になる**（作ったときに実際に起きた）
    doc = LH.document_fromstring(html.read_text(encoding="utf-8"))
    for bad in doc.xpath("//script|//style"):
        bad.getparent().remove(bad)
    sls = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]") or [doc]
    out = []
    for i, sl in enumerate(sls, 1):
        t = " ".join(sl.itertext()).replace(" ", " ")
        out.append((i, re.sub(r"\s+", " ", t)))
    return out


def ml_lines(html: Path):
    """[(枚番号, [ML の各行])]。ML はリード帯（class に lead を含む枠）の段落。"""
    doc = LH.document_fromstring(html.read_text(encoding="utf-8"))
    sls = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]") or [doc]
    out = []
    for i, sl in enumerate(sls, 1):
        bands = sl.xpath(".//*[contains(@class, 'lead')]")
        if not bands:
            continue                 # 表紙・目次など、リード帯の無い枚は見ない
        band = bands[0]
        ps = band.xpath(".//p") or [band]
        lines = [re.sub(r"\s+", " ", " ".join(x.itertext()).replace("\u00a0", " ")).strip()
                 for x in ps]
        out.append((i, [x for x in lines if x]))
    return out


def find(text: str, word: str, numeric: bool):
    pat = re.escape(word)
    if numeric:
        pat = rf"(?<![\d.]){pat}(?![\d])"
    return [m.start() for m in re.finditer(pat, text)]


def context(text: str, pos: int, word: str) -> str:
    a, b = max(0, pos - 12), min(len(text), pos + len(word) + 12)
    return text[a:b].strip()


def main():
    ap = argparse.ArgumentParser(
        description="台帳で照合する。古い語・古い値が0件かを確かめる。直した後、報告の前に使う")
    ap.add_argument("html", nargs="+", help="資料の HTML")
    ap.add_argument("--terms", help="用語台帳（terms.md）")
    ap.add_argument("--numbers", help="数値台帳（numbers.md）")
    ap.add_argument("--ml", action="store_true",
                    help="メッセージラインの型を見る（2行・各行40字以内。1行目＝何をするか、2行目＝結論）")
    ap.add_argument("--history", action="store_true",
                    help="修正の経緯が漏れた印（〜ではなく・まだ・前回 など）も出す")
    args = ap.parse_args()

    checks = []            # (種類, 語, 数値か, 落とすか)
    if args.terms:
        for r in table_rows(Path(args.terms)):
            for w in split_words(r.get("使わない語", "")):
                checks.append(("使わない語", w, False, True))
    if args.numbers:
        for r in table_rows(Path(args.numbers)):
            for w in split_words(r.get("以前の値", "")):
                checks.append(("以前の値", w, True, True))
    if args.history:
        for w in HISTORY:
            checks.append(("経緯の印", w, False, False))
    if not checks and not args.ml:
        sys.exit("照合するものが無い。--terms / --numbers / --history / --ml のどれかを渡す")

    hits, n_slides = [], 0
    for f in args.html:
        for no, text in slides_text(Path(f)):
            n_slides += 1
            for kind, w, num, fatal in checks:
                for pos in find(text, w, num):
                    hits.append((Path(f).name, no, kind, w, context(text, pos, w), fatal))

    # **ML の型。**書いてあっても守られなかったので、機械で見る
    if args.ml:
        for f in args.html:
            for no, lines in ml_lines(Path(f)):
                n_slides += 0
                if len(lines) != 2:
                    hits.append((Path(f).name, no, "MLの型",
                                 f"{len(lines)}行", " / ".join(lines)[:30], True))
                for ln in lines:
                    if len(ln) > 40:
                        hits.append((Path(f).name, no, "MLの長さ",
                                     f"{len(ln)}字", ln[:30], True))

    fatal = [h for h in hits if h[5]]
    soft = [h for h in hits if not h[5]]
    if hits:
        print("| ファイル | 枚 | 種類 | 語 | 前後 |")
        print("|---|---|---|---|---|")
        for h in fatal + soft:
            print(f"| {h[0]} | {h[1]} | {h[2]} | {h[3]} | …{h[4]}… |")
    if fatal:
        print(f"\n**古い語・古い値が {len(fatal)} 件残っている。**直してから報告する。")
        print("**置換で直さない。**その枚を作り直す（数値は同じ数字が別の項目にも出る）。")
    elif args.terms or args.numbers:
        print("\n**古い語・古い値は 0 件。**")
    if soft:
        print(f"\n修正の経緯が漏れた印が {len(soft)} 件。**前の版を知らない人が読んで通じるか**で見直す"
              "（`40_japanese.md` 3-3）。**打ち消す相手が資料の中にあれば、そのままでよい。**")

    code = checked.summary("ledger_check", n_slides, "枚", len(fatal),
                           {"照合語": len(checks), "経緯の印": len(soft)})
    return code or (1 if fatal else 0)


if __name__ == "__main__":
    sys.exit(main())

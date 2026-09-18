#!/usr/bin/env python3
"""段階0 — 1枚だけ pptx にして、テンプレート自体のずれを潰す。

    python3 scripts/stage0.py <テンプレ>/003.html -o work/stage0

**全ページ書いてから見つけると、全ページ手戻りになる。**
テンプレート由来のずれは全スライドに効くので、最初に1枚で当たりを取る。

実績: リード帯の縦位置が HTML は上寄せ、pptx は中央だった。
1枚で当たりを取っていれば、そこで潰せた。

やること。

1. HTML を PNG にする（利用者が見る絵）
2. 同じ HTML を pptx にして PNG にする（提出先が見る絵）
3. レイアウトを検査する
4. **2枚を並べて、自分の目で見比べる**

**ここで一致しない項目は、テンプレートの直しどころ。** 資料の中身ではない。
"""

import argparse
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

CHECKLIST = """
## 見比べる順（ずれやすい順）

- [ ] スライドサイズ（pt が元 PPTX と同じか）
- [ ] 共通パーツの位置（ロゴ・フッター帯・ページ番号・区切り線）
- [ ] **文字の縦位置**（帯や枠の中で上寄せか中央か）
- [ ] 行間と折り返し位置
- [ ] 書体（**手元に無い書体は代替される。そのときは判定できない**）
- [ ] 色

**ずれを見つけたら直すのは HTML テンプレート。** 出力された pptx を手で直さない。
次の変換で消えるうえ、テンプレートとして配ったときに同じずれが再発する。
"""


def run(cmd, quiet=False):
    r = subprocess.run(cmd, capture_output=True, text=True)
    if not quiet:
        print(r.stdout.rstrip())
        if r.returncode not in (0, 1) and r.stderr:
            print(r.stderr.rstrip(), file=sys.stderr)
    return r.returncode


def main():
    ap = argparse.ArgumentParser(description="段階0: 1枚でテンプレートのずれを潰す。変換の前に1枚だけ。本番の変換は html2pptx.py")
    ap.add_argument("html", help="テンプレートの1枚（本文ページが向く。例: 003.html）")
    ap.add_argument("-o", "--out", default="work/stage0", help="出力先")
    ap.add_argument("--slide", type=int, default=1,
                    help="何枚目を見るか（既定 1）。**本文ページが向く**")
    args = ap.parse_args()

    src = Path(args.html)
    if not src.exists():
        sys.exit(f"見つからない: {src}")
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    # **1枚だけにする。**受け取る deck.html は `.slide` を何枚も持つので、
    # そのまま渡すと全部変換され、「1枚で当たりを取る」意味が消える
    # （実案件の 18 枚入りを渡したところ、18 枚とも変換された）。
    import lxml.html as LH
    doc = LH.parse(str(src)).getroot()
    sls = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]")
    if len(sls) > 1:
        keep = sls[min(args.slide, len(sls)) - 1]
        for other in sls:
            if other is not keep:
                other.getparent().remove(other)
        # **元の HTML と同じ場所に書く。**別の場所に書くと
        # `images/logo.png` のような相対の参照が切れ、ロゴが落ちる
        one = src.parent / f"{src.stem}_1枚.html"
        one.write_bytes(LH.tostring(doc, encoding="utf-8"))
        print(f"**{len(sls)} 枚のうち {args.slide} 枚目だけを見る。** → {one.name}\n")
        src = one
    pptx = out / f"{src.stem}_stage0.pptx"

    print(f"# 段階0 — {src}\n")
    print("## 1. HTML の絵（利用者が見るもの）")
    rc1 = run([sys.executable, str(HERE / "html_png.py"), str(src),
               "-o", str(out / "html")])
    print("\n## 2. pptx にする")
    run([sys.executable, str(HERE / "html2pptx.py"), str(src),
         "-o", str(pptx), "--force"])
    print("\n## 3. pptx の絵（提出先が見るもの）")
    rc3 = run([sys.executable, str(HERE / "qa_render.py"), str(pptx),
               "-o", str(out / "pptx")])
    print("\n## 4. レイアウトの検査")
    rc4 = run([sys.executable, str(HERE / "fit_check.py"), "check", str(pptx)])

    print(CHECKLIST)
    print(f"HTML の絵: {out / 'html'}/")
    print(f"pptx の絵: {out / 'pptx'}/")

    if 2 in (rc1, rc3, rc4):
        print("\n**どれかが「対象0」で終わった。何も見ていない。**"
              "「ずれなし」と報告しない。")
        return 2
    print("\n**2枚を並べて自分の目で見る。** ここを飛ばすと段階0の意味がない。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""HTML を1枚1枚 PNG にする。手順5（見た目の検査）で使う。

    python3 scripts/html_png.py <テンプレ>/print.html -o qa_html
    python3 scripts/html_png.py work/p01.html work/p02.html -o qa_html

**HTML の見た目を機械が見るための唯一の手段。**
テキストと座標を読むだけでは、重なり・はみ出し・縦位置のずれは分からない。

**pptx の PNG（qa_render.py）と対にして使う。** 片方だけでは足りない。

| 見るもの | 出す道具 | そこでしか見つからないもの |
|---|---|---|
| HTML | **これ** | 利用者がブラウザで見ている絵そのもの |
| pptx | `qa_render.py` | 書体・折り返し・最終成果物の姿 |
| 両方を並べる | 両方 | **HTML と pptx のずれ**（縦位置・行間・自動調整） |

実績: リード帯の縦位置が HTML は上寄せ、pptx は中央だった。
**どちらか片方だけ見ていては出ない。** 承認した絵と出てくる絵が違う。

外に出さない。ローカルに PNG を書くだけで、公開もアップロードもしない。
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import re

from lxml import html as LH

import checked

# 探す順は OS で変える。**Windows は Edge を先に見る。**
# Edge は標準で入っており、Chrome は入っていても
# 企業の設定で起動できないことがある
_MAC = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
)
_WIN = (
    # **PATH に載っていないので絶対パスで探す**
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
)
_PATH = ("msedge", "google-chrome", "chrome", "chromium", "chromium-browser")


def candidates():
    import platform
    if platform.system() == "Windows":
        return _WIN + _PATH
    if platform.system() == "Darwin":
        return _MAC + _PATH
    return _PATH


def browser_name(path: str) -> str:
    """実際に使うブラウザの名前。**「Chrome」と決めつけない。**

    Edge で動かしているのに Chrome と表示すると、
    利用者が「Chrome を入れないといけないのか」と読む。
    """
    n = os.path.basename(path).lower()
    if "edge" in n:
        return "Microsoft Edge"
    if "chromium" in n:
        return "Chromium"
    if "chrome" in n:
        return "Google Chrome"
    return os.path.basename(path)


def find_browser(explicit=None, quiet=False):
    """使えるブラウザを探す。**見つからないことと、問題が無いことを混同しない。**"""
    if explicit:
        p = explicit if os.path.isabs(explicit) else shutil.which(explicit)
        if p and os.path.exists(p):
            return p
        if not quiet:
            print(f"指定されたブラウザが無い: {explicit}", file=sys.stderr)
        return None
    for c in candidates():
        p = c if os.path.isabs(c) else shutil.which(c)
        if p and os.path.exists(p):
            return p
    if not quiet:
        print("ブラウザが見つからない（Chromium 系が要る）。", file=sys.stderr)
        print("  Windows: **Edge は標準で入っている。**"
              "入っているのに見つからないなら --browser でパスを渡す", file=sys.stderr)
        print("  macOS:   Google Chrome か Microsoft Edge を入れる", file=sys.stderr)
        print("  Linux:   apt install chromium", file=sys.stderr)
    return None


def expand(paths):
    """print.html を渡されたら、中の iframe の並びに展開する。"""
    out = []
    for h in paths:
        p = Path(h)
        if not p.exists():
            sys.exit(f"見つからない: {p}")
        doc = LH.parse(str(p)).getroot()
        frames = doc.xpath("//iframe/@src")
        out += [p.parent / f for f in frames] if frames else [p]
    return out


SLIDE_SIZE_RE = re.compile(
    r"\.slide\s*\{[^}]*?width:\s*([\d.]+)px[^}]*?height:\s*([\d.]+)px", re.S)


def size_of(path: Path):
    """.slide の px 寸法。無ければ 1280x720 とみなす。

    **ここで CSS を直接読む。** 変換側（slider-deck の `html_abs.py`）に
    依存させない。絵を出すだけのために変換器一式を持ち込むことになる。
    """
    try:
        m = SLIDE_SIZE_RE.search(path.read_text(encoding="utf-8"))
        if m:
            return int(round(float(m.group(1)))), int(round(float(m.group(2))))
    except OSError:
        pass
    return 1280, 720


def _run(browser, src: Path, out: Path, w: int, h: int, scale: float, profile,
         headless="--headless=new"):
    """1回だけ起動してみる。

    **ヘッドレスの方式を引数にする。** 近年の Edge / Chrome は
    旧 `--headless` を落としており、**起動はするのに PNG が1枚も書かれない。**
    """
    cmd = [
        browser, headless, "--disable-gpu", "--hide-scrollbars",
        "--no-sandbox", "--force-device-scale-factor=%g" % scale,
        "--virtual-time-budget=2000",
        "--no-first-run", "--no-default-browser-check", "--disable-extensions",
        f"--window-size={w},{h}",
        # **絶対パスで渡す。**相対だとブラウザ側の作業ディレクトリを基準に
        # 書かれ、こちらの out.exists() と食い違って
        # 「PNG は出来ているのに失敗」と誤報する
        f"--screenshot={out.resolve()}",
    ]
    if profile:
        cmd.insert(1, f"--user-data-dir={profile}")
    cmd.append(src.resolve().as_uri())
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if out.exists():
        return out.stat().st_size, None
    return None, (r.stderr or r.stdout or "").strip()[-300:]


def shoot(browser, src: Path, out: Path, w: int, h: int, scale: float, profile):
    """1枚を PNG にする。**新旧のヘッドレス × プロファイル有無を順に試す。**

    | 試す順 | なぜ |
    |---|---|
    | 新ヘッドレス・素 | ふつうはこれで通る |
    | 新ヘッドレス・専用プロファイル | **ブラウザが起動中だと既定のプロファイルを掴めない** |
    | 旧ヘッドレス・素 | 古い版のブラウザ向け |
    | 旧ヘッドレス・専用プロファイル | 同上 |

    **どれかが通れば成功。全部落ちたら、全部のエラーをまとめて返す。**
    1つ目のエラーだけを見せると、本当の原因が隠れる。

    **専用プロファイルを最初から使わない。**
    macOS の Chrome では、それを付けると
    `Trying to load the allocator multiple times` で起動に失敗する（実測）。
    """
    out.parent.mkdir(parents=True, exist_ok=True)
    errs = []
    for headless in ("--headless=new", "--headless"):
        for prof in (None, profile):
            size, err = _run(browser, src, out, w, h, scale, prof, headless)
            if size is not None:
                return size, None
            errs.append(f"[{headless}{'' if prof is None else ' +profile'}] {err}")
    return None, " ／ ".join(e for e in errs if e)[-300:]


def main():
    ap = argparse.ArgumentParser(
        description="HTML を1枚1枚 PNG にする（ローカルのみ。公開しない）")
    ap.add_argument("html", nargs="*",
                    help="スライドの HTML。print.html を渡すと中の iframe 順に展開する")
    ap.add_argument("-o", "--out", default="qa_html", help="出力先ディレクトリ")
    ap.add_argument("--scale", type=float, default=1.0,
                    help="拡大率。既定 1.0。文字を細かく見たいときに 2 にする")
    ap.add_argument("--probe", action="store_true",
                    help="使える道具があるかだけ調べる")
    ap.add_argument("--browser",
                    help="使うブラウザの実行ファイル。**既定の場所に無いときに渡す**")
    args = ap.parse_args()

    browser = find_browser(args.browser, quiet=args.probe)
    if args.probe:
        print(f"ブラウザ: {browser_name(browser) if browser else '**無し**'}"
              f"{'  ' + browser if browser else ''}")
        print("無いときは見た目の検査ができない。**検査していないと明記する。**")
        return 0 if browser else 1
    if not browser:
        return 1
    if not args.html:
        ap.error("HTML を1つ以上渡す")

    files = expand(args.html)
    outdir = Path(args.out)
    rows, ng = [], 0
    # **専用のプロファイル。**起動中のブラウザと喧嘩しないため（shoot を見る）
    with tempfile.TemporaryDirectory(prefix="html_png-") as profile:
        for i, f in enumerate(files, 1):
            w, h = size_of(f)
            png = outdir / f"{i:02d}_{f.stem}.png"
            size, err = shoot(browser, f, png, w, h, args.scale, profile)
            if err:
                ng += 1
                rows.append((i, f.name, f"{w}×{h}", "**失敗**", err[:60]))
            else:
                rows.append((i, f.name, f"{w}×{h}", png.name, f"{size/1024:.0f}KB"))

    print(f"HTML → PNG: {len(files)} 枚 → {outdir}/（失敗 {ng} 件）")
    print("| # | HTML | px | PNG | |")
    print("|---|---|---|---|---|")
    for r in rows:
        print("| " + " | ".join(str(x) for x in r) + " |")
    if ng:
        print("\n**失敗した枚は見ていない。**「問題なし」と報告しない。")
        print("  ブラウザが起動中でも動くように専用プロファイルを使っている。"
              "それでも失敗するなら、**`--browser` で実行ファイルを明示する。**")
        return 1
    print("\n**1枚ずつ自分で見る。** 抜き取りでは見落とす。")
    print("**pptx の PNG とも並べる**（qa_render.py）。ずれは片方だけでは出ない。")
    return checked.summary("html_png", len(files) - ng, "枚", ng, {
        "描画": browser_name(browser),
        "拡大率": args.scale,
        "書体": checked.font_state("verdana", "meiryo"),
    })


if __name__ == "__main__":
    sys.exit(main())

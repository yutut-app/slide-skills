#!/usr/bin/env python3
"""pptx / xlsx を1ページ1枚の PNG にする。手順5（見た目の検査）で使う。

    python3 scripts/qa_render.py deck.pptx [-o 出力先]

**PDF を経由しない。** LibreOffice の PDF 出力は、端末に無いフォントを
日本語グリフを持たないフォント（Noto Sans など）に置き換えて埋め込むため、
日本語が文字化けする。直接 PNG に変換する経路ではラスタライズ時に
フォールバックが効き、日本語が正しく出る。

LibreOffice の PNG 出力は先頭ページしか出さないので、pptx は
`<p:sldIdLst>` を1枚ぶんに絞った一時ファイルをスライドの数だけ作って変換する。

出力は <出力先>/<元ファイル名>-01.png ... 。既定の出力先は ./qa 。
生成したファイルのパスを標準出力に1行ずつ出す。
**必ず画像を開いて自分で見る。** 生成できたことは、正しさの証拠にならない。

最後に `検査: qa_render 対象 n枚 指摘 n件 前提: ...` の1行を出す。
**1枚も出せなかったときは異常終了する**（`checked.py`）。
0枚と「問題なし」が出力で区別できないと、検査を飛ばしたことに気づけない。
"""

import argparse
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked

P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
]


def find_soffice(quiet=False):
    """LibreOffice は PATH に無いことが多いので実体を順に探す。"""
    for p in SOFFICE_CANDIDATES:
        if Path(p).exists():
            return p
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    if quiet:
        return None
    sys.exit(
        "LibreOffice が見つからない。\n"
        "  brew install --cask libreoffice\n"
        "入れられない場合は、見た目の検査を行っていないことを明記して"
        "利用者に目視を依頼する。黙って飛ばさない。"
    )


def convert_png(soffice, src: Path, outdir: Path) -> Path:
    subprocess.run(
        [soffice, "--headless", "--convert-to", "png", "--outdir", str(outdir), str(src)],
        check=True, capture_output=True,
    )
    png = outdir / (src.stem + ".png")
    if not png.exists():
        sys.exit(f"PNG に変換できなかった: {src}")
    return png


def slide_ids(pptx: Path):
    """<p:sldIdLst> の子要素を順に返す。要素そのものではなく XML 断片で持つ。"""
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(pptx) as z:
        xml = z.read("ppt/presentation.xml").decode("utf-8")
    root = ET.fromstring(xml)
    lst = root.find(f"{P_NS}sldIdLst")
    if lst is None:
        return xml, []
    return xml, list(lst)


def single_slide_copy(pptx: Path, index: int, dest: Path):
    """スライド index（0始まり）だけを <p:sldIdLst> に残した pptx を作る。

    スライドのファイル自体は消さない。一覧から外れたものは描画されないので、
    これだけで1枚ぶんの PNG が得られる。ファイルを消すと関連付けが壊れる。
    """
    import xml.etree.ElementTree as ET

    ET.register_namespace(
        "p", "http://schemas.openxmlformats.org/presentationml/2006/main")
    ET.register_namespace(
        "a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    ET.register_namespace(
        "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    with zipfile.ZipFile(pptx) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}

    root = ET.fromstring(blobs["ppt/presentation.xml"].decode("utf-8"))
    lst = root.find(f"{P_NS}sldIdLst")
    keep = list(lst)[index]
    for child in list(lst):
        if child is not keep:
            lst.remove(child)
    blobs["ppt/presentation.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, blobs[n])



def probe():
    """使える画像化の手段を先に確認する。**環境で違う。**

    無いものを前提にした手順を書くと、その環境では検査が飛ばされる。
    """
    import platform
    print(f"OS: {platform.system()} {platform.release()}")
    found = []
    p = find_soffice(quiet=True)
    if p:
        print(f"  LibreOffice: {p}")
        found.append("soffice")
    else:
        print("  LibreOffice: 無し")
    if platform.system() == "Windows":
        try:
            import win32com.client  # noqa: F401
            print("  PowerPoint COM: 使える（pywin32 あり）")
            found.append("com")
        except ImportError:
            print("  PowerPoint COM: pywin32 が無い（pip install pywin32）")
    if not found:
        print("\n**画像化の手段が無い。**")
        print("検査していないことを明記して、利用者に目視を依頼する。黙って飛ばさない。")
        return 1
    print(f"\n使う手段: {found[0]}")
    return 0


def render_windows_com(src: Path, outdir: Path) -> list:
    """Windows で LibreOffice が無いときの経路。PowerPoint 本体に書き出させる。

    ppSaveAsPNG = 18。SaveCopyAs で元ファイルを触らずに書き出す。
    スライドごとに連番の PNG が入ったフォルダができる。
    """
    import win32com.client

    outdir.mkdir(parents=True, exist_ok=True)
    dest = outdir / src.stem
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = None
    try:
        pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
        pres.SaveCopyAs(str(dest.resolve()), 18)  # 18 = ppSaveAsPNG
    finally:
        if pres is not None:
            pres.Close()
        app.Quit()
    return sorted(dest.glob("*.PNG")) + sorted(dest.glob("*.png"))


def render(src: Path, outdir: Path) -> list[Path]:
    soffice = find_soffice(quiet=True)
    if soffice is None:
        import platform
        if platform.system() == "Windows":
            return render_windows_com(src, outdir)
        find_soffice()  # ここで案内を出して終了する
    
    outdir.mkdir(parents=True, exist_ok=True)
    made = []

    if src.suffix.lower() not in (".pptx", ".potx"):
        # xlsx などは1ページぶんだけ出る。表とグラフが1ページに収まる設定が前提
        with tempfile.TemporaryDirectory() as tmp:
            png = convert_png(soffice, src, Path(tmp))
            dest = outdir / f"{src.stem}-01.png"
            shutil.copy(png, dest)
            made.append(dest)
        return made

    _, ids = slide_ids(src)
    if not ids:
        sys.exit(f"スライドが見つからない: {src}")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        for i in range(len(ids)):
            one = tmpdir / f"{src.stem}_s{i + 1:02d}.pptx"
            single_slide_copy(src, i, one)
            png = convert_png(soffice, one, tmpdir)
            dest = outdir / f"{src.stem}-{i + 1:02d}.png"
            shutil.copy(png, dest)
            made.append(dest)
    return made


def means():
    """実際に使える画像化の手段。**前提欄に書くため。**"""
    import platform
    if find_soffice(quiet=True):
        return "LibreOffice"
    if platform.system() == "Windows":
        try:
            import win32com.client  # noqa: F401
            return "PowerPoint COM"
        except ImportError:
            pass
    return "**無し**"


def main():
    ap = argparse.ArgumentParser(description="pptx/xlsx を PNG にして目視検査する。自分の目で見るときに使う")
    ap.add_argument("files", nargs="*", help="検査する pptx / xlsx")
    ap.add_argument("-o", "--outdir", default="qa", help="出力先（既定: ./qa）")
    ap.add_argument("--probe", action="store_true",
                    help="使える画像化の手段を確認するだけ。**検査の前に必ず一度実行する**")
    args = ap.parse_args()

    if args.probe:
        sys.exit(probe())

    made, missing = 0, 0
    for f in args.files:
        src = Path(f)
        if not src.exists():
            print(f"見つからない: {src}", file=sys.stderr)
            missing += 1
            continue
        for p in render(src, Path(args.outdir)):
            print(p)
            made += 1

    if made:
        print("\n**1枚ずつ開いて自分で見る。** 生成できたことは、正しさの証拠にならない。")
        print("**HTML の絵とも並べる**（scripts/html_png.py）。"
              "ずれは片方だけでは出ない。")
    return checked.summary("qa_render", made, "枚", missing, {
        "手段": means(),
        "書体": checked.font_state("verdana", "meiryo"),
        "出力先": args.outdir,
    })


if __name__ == "__main__":
    sys.exit(main())

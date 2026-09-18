#!/usr/bin/env python3
"""HTML の図をもとに **Excel のブックを作り、その中のグラフを pptx に貼る。**

    # 1. HTML から .xlsx を起こす（どの環境でも動く）
    python3 scripts/chart_xlsx.py deck/deck.html -o deck/charts

    # 2. Excel のグラフを pptx に貼る（**Windows + Excel + PowerPoint のみ**）
    python3 scripts/chart_xlsx.py --paste deck/charts deck/deck_rev1.pptx

## なぜ Excel を経由するか

提出先で**数字を Excel で直す**運用のため。ネイティブのグラフも中に
ワークシートを持つが、**別ファイルの .xlsx なら Excel でそのまま開ける。**

## 数はどこから取るか

**推測しない。**数が読めないものは作らずに報告する。

| 出どころ | 優先 |
|---|---|
| `data-values` / `data-labels`（作る側が書いたもの） | 1 |
| 図の中の文字に並んでいる「ラベル 数値」の組 | 2 |
| **どちらも無い** | **作らない。**配布元に数を足してもらう |

貼り付けは COM（PowerPoint/Excel 本体）で行う。**本体が要る。**
macOS では 1 だけを使い、貼り付けは Windows 側で行う。
"""

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
import lxml.html as LH

KIND_TO_XL = {
    "column": ("bar", "col"), "bar": ("bar", "bar"),
    "column-stacked": ("bar", "col"), "bar-stacked": ("bar", "bar"),
    "line": ("line", None), "line-plain": ("line", None),
    "pie": ("pie", None), "doughnut": ("pie", None),
    "area": ("area", None), "scatter": ("scatter", None),
}

_NUM = re.compile(r"-?[\d,]+(?:\.\d+)?")


def _numbers(s):
    """文中から数を拾う。**桁区切りのカンマだけを外す。**"""
    return [float(t.replace(",", "")) for t in _NUM.findall(s or "")]


def _values(s):
    """`data-values` の並び。**区切りのカンマを桁区切りと読み違えない。**

    実績: `62,18,11` を1つの数（62181175…）として読み、表が壊れた。
    区切りで切ってから数にする。
    """
    out = []
    for t in re.split(r"[,、\s]+", (s or "").strip()):
        if not t:
            continue
        try:
            out.append(float(t))
        except ValueError:
            return []
    return out


def read_series(el):
    """(種別, ラベル, [(系列名, 値)]) を返す。読めなければ None。"""
    kind = (el.get("data-chart") or "").strip().lower()
    labels = [t.strip() for t in (el.get("data-labels") or "").split(",") if t.strip()]
    series = []
    for i in range(1, 9):
        sfx = "" if i == 1 else f"-{i}"
        vals = _values(el.get("data-values" + sfx))
        if vals:
            series.append(((el.get("data-series" + sfx) or f"系列{i}").strip(), vals))
    if series:
        if not labels:
            labels = [str(i + 1) for i in range(len(series[0][1]))]
        return kind or "column", labels, series

    # **書かれていないときは、図の中の文字から拾う。**
    # 「ラベル 数値」が縦に並んでいる形だけを見る。読めなければ作らない
    pairs = []
    for node in el.iter():
        if not isinstance(node.tag, str):
            continue
        t = " ".join((node.text or "").split())
        if not t:
            continue
        nums = _numbers(t)
        if len(nums) == 1:
            name = _NUM.sub("", t).strip(" 　:：・")
            if name:
                pairs.append((name, nums[0]))
    if len(pairs) >= 3:
        return (kind or "column", [p[0] for p in pairs],
                [("値", [p[1] for p in pairs])])
    return None


def write_xlsx(path: Path, title, kind, labels, series):
    """データの表と、その表を参照するグラフを1枚に入れる。

    **グラフは表を参照する。**値を直接書き込むと、Excel で数字を直しても
    グラフが変わらない。
    """
    from openpyxl import Workbook
    from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference

    wb = Workbook()
    ws = wb.active
    ws.title = "データ"
    ws["A1"] = title
    ws.append([])
    ws.append(["分類"] + [name for name, _ in series])
    for i, lab in enumerate(labels):
        ws.append([lab] + [(vals[i] if i < len(vals) else None) for _, vals in series])

    base, sub = KIND_TO_XL.get(kind, ("bar", "col"))
    ch = {"bar": BarChart, "line": LineChart,
          "pie": PieChart, "area": AreaChart}.get(base, BarChart)()
    if base == "bar":
        ch.type = sub or "col"
        if "stacked" in kind:
            ch.grouping = "stacked"
            ch.overlap = 100
    ch.title = title
    last = 3 + len(labels)
    ch.add_data(Reference(ws, min_col=2, max_col=1 + len(series), min_row=3,
                          max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=last))
    ch.height, ch.width = 8, 14
    ws.add_chart(ch, f"A{last + 2}")
    ws.column_dimensions["A"].width = 22
    wb.save(path)


def box_px(el, rules, inherit):
    """図の枠を px で返す。**取れないものは None。**推測で埋めない。"""
    import html_abs as HA
    st = HA.computed(el, rules, inherit)
    while st.get("left") is None and el.getparent() is not None:
        el = el.getparent()
        st = HA.computed(el, rules, inherit)
    return {k: HA.px(st.get(k)) for k in ("left", "top", "width", "height")}


def extract(html: Path, outdir: Path, warn):
    """HTML の中の図を1つずつ .xlsx にする。作れた数を返す。"""
    import html_abs as HA

    doc = LH.parse(str(html)).getroot()
    rules = []
    for st in doc.xpath("//style"):
        rules += HA.parse_css(st.text or "")
    body = doc.find(".//body")
    inherit = HA.inheritable(HA.computed(body, rules) if body is not None else {})

    slides = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]") or [doc]
    made, skipped, geom = 0, 0, {}
    outdir.mkdir(parents=True, exist_ok=True)
    # **(係数, 出どころ) で返る。**係数だけを控える
    px_per_pt, px_src = HA.px_per_pt_of([html])
    for sno, slide in enumerate(slides, 1):
        cands = slide.xpath(".//*[@data-chart]") + slide.xpath(".//svg") + \
            slide.xpath(".//*[@data-figure]")
        # **箱で描いたグラフも拾う。**実際の資料は棒が <div> のことがある。
        # ここで拾わないと、変換だけが警告を出して、作る道具は空振りする
        for el in slide.iterchildren():
            if isinstance(el.tag, str) and el.tag not in ("svg", "canvas") \
                    and HA.looks_like_chart(el, rules, inherit):
                cands.append(el)
        seen = set()
        for el in cands:
            if id(el) in seen:
                continue
            seen.add(id(el))
            got = read_series(el)
            name = (el.get("data-figure") or el.findtext("{*}title")
                    or f"図{len(seen)}").strip()
            if got is None:
                warn.append(f"スライド{sno}「{name}」: **数が読めない。**"
                            "`data-values` を配布元に足してもらう")
                skipped += 1
                continue
            kind, labels, series = got
            out = outdir / f"s{sno:02d}_{re.sub(r'[^0-9A-Za-zぁ-んァ-ヶ一-龠]', '_', name)[:20]}.xlsx"
            write_xlsx(out, name, kind, labels, series)
            # **貼る位置を控えておく。**貼り付けは中央に入るので、
            # 元の図の場所が分からないと、人が毎回動かすことになる
            geom[out.name] = {"slide": sno, "name": name,
                              **box_px(el, rules, inherit)}
            print(f"  {out.name}  {kind} / 分類{len(labels)} / 系列{len(series)}")
            made += 1
    if geom:
        import json
        (outdir / "geometry.json").write_text(
            json.dumps({"px_per_pt": px_per_pt, "係数の出どころ": px_src,
                        "図": geom},
                       ensure_ascii=False, indent=2), encoding="utf-8")
    return made, skipped


def paste(xdir: Path, pptx: Path, warn, dry=False):
    """Excel のグラフを PowerPoint に貼る。**Windows の本体が要る。**

    貼り先のスライドは、ファイル名の先頭 `sNN_` で決める。
    **貼るのは図形としてのグラフ**（リンクではない）。リンクにすると、
    受け取った側で .xlsx の場所が変わったときに開けなくなる。
    """
    import json
    import platform

    g = {}
    gf = xdir / "geometry.json"
    if gf.exists():
        g = json.loads(gf.read_text(encoding="utf-8"))

    if dry:
        return dry_run(xdir, pptx, g, warn)

    if platform.system() != "Windows":
        warn.append("**貼り付けは Windows でしか行えない**"
                    "（Excel と PowerPoint の本体を使う）。"
                    ".xlsx は作ってあるので、Windows 側で `--paste` を実行する")
        return 0
    import win32com.client as w

    files = sorted(xdir.glob("*.xlsx"))
    if not files:
        return 0
    xl = w.Dispatch("Excel.Application")
    pp = w.Dispatch("PowerPoint.Application")
    pres = pp.Presentations.Open(str(pptx.resolve()), WithWindow=False)
    n = 0
    try:
        for f in files:
            m = re.match(r"s(\d+)_", f.name)
            sno = int(m.group(1)) if m else 1
            if sno > pres.Slides.Count:
                warn.append(f"{f.name}: スライド{sno}が無い（全{pres.Slides.Count}枚）")
                continue
            wb = xl.Workbooks.Open(str(f.resolve()))
            try:
                ch = find_chart(wb)
                if ch is None:
                    warn.append(f"{f.name}: **グラフが見つからない**")
                    continue
                ch.ChartArea.Copy()
                # 2 = ppPasteShape。**貼り付け先の体裁に合わせない**
                shp = pres.Slides(sno).Shapes.PasteSpecial(2)(1)
                place(shp, g, f.name, warn)
                n += 1
            finally:
                wb.Close(False)
        pres.Save()
    finally:
        pres.Close()
        xl.Quit()
        pp.Quit()
    return n


def find_chart(wb):
    """ブックの中の最初のグラフ。**1枚目に無いことがある。**全シート見る。"""
    for i in range(1, wb.Sheets.Count + 1):
        sh = wb.Sheets(i)
        if sh.ChartObjects().Count:
            return sh.ChartObjects(1).Chart
    return None


def place(shp, g, name, warn):
    """控えた px の位置に置く。**控えが無ければ動かさない。**

    勝手に中央から動かすより、**動かしていないと分かるほうがよい。**
    """
    info = (g.get("図") or {}).get(name)
    ppp = g.get("px_per_pt")
    if not info or not ppp or info.get("left") is None:
        warn.append(f"{name}: 位置の控えが無い。**中央に入ったまま。**手で置く")
        return
    shp.Left = info["left"] / ppp
    shp.Top = info["top"] / ppp
    if info.get("width"):
        shp.Width = info["width"] / ppp
    if info.get("height"):
        shp.Height = info["height"] / ppp


def dry_run(xdir: Path, pptx: Path, g, warn):
    """**Windows へ持ち込む前に、通らない条件をここで出す。**

    COM が無い環境でも、貼り先の枚数・位置の控え・対応づけは確かめられる。
    """
    from pptx import Presentation

    files = sorted(xdir.glob("*.xlsx"))
    n_slides = len(Presentation(str(pptx)).slides)
    print(f"貼り先: {pptx.name}（{n_slides}枚）  ブック {len(files)} 件\n")
    print("| ブック | スライド | 位置(pt) |")
    print("|---|---|---|")
    ok = 0
    for f in files:
        m = re.match(r"s(\d+)_", f.name)
        sno = int(m.group(1)) if m else 1
        info = (g.get("図") or {}).get(f.name)
        ppp = g.get("px_per_pt")
        if sno > n_slides:
            pos = "**貼り先のスライドが無い**"
            warn.append(f"{f.name}: スライド{sno}が無い（全{n_slides}枚）")
        elif not info or info.get("left") is None or not ppp:
            pos = "**控えが無い → 中央**"
            warn.append(f"{f.name}: 位置の控えが無い")
        else:
            pos = (f"{info['left'] / ppp:.0f},{info['top'] / ppp:.0f} "
                   f"{(info.get('width') or 0) / ppp:.0f}×"
                   f"{(info.get('height') or 0) / ppp:.0f}")
            ok += 1
        print(f"| {f.name} | {sno} | {pos} |")
    return ok


def main():
    ap = argparse.ArgumentParser(description="HTML の図から Excel を作り、pptx に貼る。Excel でグラフを作って貼るときに使う")
    ap.add_argument("src", help="HTML（既定）／--paste のときは .xlsx のあるフォルダ")
    ap.add_argument("pptx", nargs="?", help="--paste のときの貼り先")
    ap.add_argument("-o", "--outdir", default="charts", help="xlsx の出力先")
    ap.add_argument("--paste", action="store_true",
                    help="作った Excel のグラフを pptx に貼る（**Windows のみ**）")
    ap.add_argument("--dry-run", action="store_true",
                    help="**貼らずに対応づけだけ確かめる。**どの環境でも動く。"
                         "Windows へ持ち込む前にこれを通す")
    args = ap.parse_args()

    warn = []
    if args.paste:
        if not args.pptx:
            sys.exit("貼り先の pptx を渡す: --paste <xlsxのフォルダ> <pptx>")
        n = paste(Path(args.src), Path(args.pptx), warn, dry=args.dry_run)
        for w_ in warn:
            print("- " + w_)
        return checked.summary(
            "chart_xlsx --paste" + (" --dry-run" if args.dry_run else ""),
            n, "図", len(warn), {"貼り先": args.pptx})

    made, skipped = extract(Path(args.src), Path(args.outdir), warn)
    for w_ in warn:
        print("- " + w_)
    if made:
        print(f"\n**Excel で開いて数字を確かめる。**{args.outdir}/")
        print("貼り付けは Windows 側: "
              f"python3 scripts/chart_xlsx.py --paste {args.outdir} <pptx>")
    return checked.summary("chart_xlsx", made, "図", skipped,
                           {"数が読めなかった図": skipped})


if __name__ == "__main__":
    sys.exit(main())

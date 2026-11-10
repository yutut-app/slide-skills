#!/usr/bin/env python3
"""グラフテンプレート（Excel）を生成する。

1テンプレ = 1ファイル = 1シート。
各シートは「データ表」「ネイティブグラフ」「使い方メモ」の3点で構成する。

使い方:
    python3 scripts/gen_chart_templates.py [出力先ディレクトリ]
    既定の出力先は assets/templates/charts
"""

import re
import sys
import pathlib
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    BarChart,
    LineChart,
    PieChart,
    RadarChart,
    Reference,
    ScatterChart,
    Series,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.trendline import Trendline
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 配色。提案・報告資料向けに彩度を抑える。1色を主役にし、残りは補助とグレー。
# 公式 pptx スキルの派手なパレットは使わない。
# ---------------------------------------------------------------------------
NAVY = "1F4E79"      # 主色。強調したい系列に使う
BLUE = "4472C4"      # 副色
LIGHT = "BDD7EE"     # 淡色。背景側の系列
GRAY = "808080"      # 比較対象・前年など主張しない系列
ACCENT = "C00000"    # 基準線・目標線・逸脱の指摘にのみ使う
INK = "333333"       # 文字

FONT = "游ゴシック"   # 日本語
LATIN = None          # 欧文。None なら日本語と同じ
FONT_FALLBACK = "Meiryo"  # 游ゴシックが無い環境向け。メモに記載する

HEAD_FILL = PatternFill("solid", fgColor=NAVY)
NOTE_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def setup_page(ws):
    """印刷・PDF化したときに表とグラフが1ページに収まるようにする。

    既定のままだと縦向き1ページ幅に収まらず、グラフだけ次ページに落ちる。
    資料に貼る前に印刷プレビューで確認されることが多いので、ここで揃えておく。
    """
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def style_sheet(ws, title, purpose, notes, data_ref, ncols):
    """シート共通の体裁。見出し・用途・使い方メモを置く。"""
    ws.sheet_view.showGridLines = False
    setup_page(ws)

    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
    ws["A2"] = purpose
    ws["A2"].font = Font(name=FONT, size=10, color=INK)

    # データ表の見出し行を塗る
    hdr = data_ref["header_row"]
    first = data_ref["first_col"]
    for c in range(first, first + ncols):
        cell = ws.cell(row=hdr, column=c)
        cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    for r in range(hdr + 1, data_ref["last_row"] + 1):
        for c in range(first, first + ncols):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=FONT, size=10, color=INK)
            cell.border = BORDER
            if c > first:
                cell.alignment = Alignment(horizontal="right")

    # 使い方メモ
    note_row = data_ref["last_row"] + 2
    ws.cell(row=note_row, column=first, value="使い方").font = Font(
        name=FONT, size=10, bold=True, color=NAVY
    )
    for i, line in enumerate(notes, start=1):
        cell = ws.cell(row=note_row + i, column=first, value=line)
        cell.font = Font(name=FONT, size=9, color=INK)
        cell.fill = NOTE_FILL
        cell.alignment = Alignment(vertical="center")

    ws.column_dimensions["A"].width = 3
    for c in range(first, first + max(ncols, 4)):
        ws.column_dimensions[get_column_letter(c)].width = 16



def chart_font(ch):
    """後方互換のための空処理。

    openpyxl の ChartBase には txPr が無く（ChartSpace 側の要素）、代入しても
    黙って無視される。書体は保存後に inject_chart_font() で入れる。
    """
    return ch


def inject_chart_font(path):
    """保存した xlsx のグラフに、既定の書体を入れる。

    グラフの中の文字（タイトル・軸・凡例・ラベル）は、指定しないとテーマの
    既定になる。**表はテンプレートの書体、グラフだけ別**という状態を避ける。
    ea を書かないと日本語に効かない点はスライドと同じ。

    <c:txPr> は <c:chart> の後ろに置く必要がある（スキーマ上の順序）。
    """
    import shutil
    import zipfile

    latin = LATIN or FONT
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    # openpyxl は接頭辞なし（既定名前空間）で書く。PowerPoint 由来だと c: が付く。
    # **閉じタグの形を実測して合わせる。**決め打ちすると黙って何も起きない
    def make(prefix):
        p = f"{prefix}:" if prefix else ""
        return (
            f'<{p}txPr><a:bodyPr xmlns:a="{A}"/><a:lstStyle xmlns:a="{A}"/>'
            f'<a:p xmlns:a="{A}"><a:pPr><a:defRPr>'
            f'<a:latin typeface="{latin}"/><a:ea typeface="{FONT}"/>'
            f'<a:cs typeface="{latin}"/>'
            f'</a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></{p}txPr>'
        )

    import shutil
    import zipfile

    src = zipfile.ZipFile(path)
    tmp = str(path) + ".tmp"
    touched = 0
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if re.match(r"xl/charts/chart\d+\.xml$", item.filename):
                x = data.decode("utf-8")
                for prefix, close in (("", "</chart>"), ("c", "</c:chart>")):
                    if close in x and "txPr>" not in x:
                        x = x.replace(close, close + make(prefix), 1)
                        data = x.encode("utf-8")
                        touched += 1
                        break
            out.writestr(item, data)
    src.close()
    shutil.move(tmp, path)
    return touched


def value_labels(ch, percent=False, cat_name=False):
    """データラベルを付ける。表示する項目を全部明示する。

    showVal だけ立てて他を未指定にすると、レンダラ既定で系列名やカテゴリ名まで
    出て「寸法不良（外径）: 件数: 186」のような読めないラベルになる。
    """
    dl = DataLabelList()
    dl.showVal = not percent
    dl.showPercent = percent
    dl.showCatName = cat_name
    dl.showSerName = False
    dl.showLegendKey = False
    dl.showBubbleSize = False
    ch.dLbls = dl
    return ch


def style_chart(ch, title, x_title=None, y_title=None, w=20, h=11):
    ch.title = title
    ch.width = w
    ch.height = h
    ch.style = None
    # 円グラフには軸が無いので、属性の有無を見てから触る
    if x_title and hasattr(ch, "x_axis"):
        ch.x_axis.title = x_title
    if y_title and hasattr(ch, "y_axis"):
        ch.y_axis.title = y_title
    return ch


def solid(series, hexcolor, line=False):
    """系列に単色を当てる。line=True なら線の色。"""
    if line:
        series.graphicalProperties.line.solidFill = hexcolor
        series.graphicalProperties.line.width = 22000  # 約1.75pt
    else:
        series.graphicalProperties.solidFill = hexcolor
        series.graphicalProperties.line.noFill = True
    return series


def write_table(ws, rows, first_row=4, first_col=2):
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            ws.cell(row=first_row + i, column=first_col + j, value=val)
    return {
        "header_row": first_row,
        "first_col": first_col,
        "last_row": first_row + len(rows) - 1,
    }


def save(wb, outdir, name):
    path = outdir / name
    wb.save(path)
    inject_chart_font(path)
    print(f"  {path.name}")
    return path


# ---------------------------------------------------------------------------
# 各テンプレート
# ---------------------------------------------------------------------------

def t_bar_vertical(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "縦棒グラフ"
    rows = [["項目", "実績"]] + [
        [f"項目{c}", v] for c, v in zip("ABCDE", [186, 152, 118, 74, 41])
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "縦棒グラフ（大小の比較）",
        "項目間の量の大小を比べる。連続していない項目に使う。",
        [
            "・項目は多くても7つまで。それ以上は横棒グラフか、上位N＋その他にまとめる",
            "・降順に並べる。並び順に意味がある場合（月・工程順）を除く",
            "・縦軸は必ず0から始める。途中から始めると差が誇張される",
            "・強調したい1本だけ濃紺、残りはグレーにすると主張が伝わる",
        ],
        ref, 2,
    )
    ch = BarChart()
    ch.type = "col"
    ch.gapWidth = 60
    data = Reference(ws, min_col=3, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    solid(ch.series[0], NAVY)
    value_labels(ch)
    ch.legend = None
    style_chart(ch, "項目別の実績", y_title="件数")
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "01_縦棒グラフ.xlsx")


def t_bar_horizontal(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "横棒グラフ"
    rows = [["項目", "実績"]] + [
        [n, v] for n, v in [
            ("寸法不良（外径）", 186), ("キズ・打痕", 152), ("組付け位置ずれ", 118),
            ("表面粗さ不良", 74), ("刻印かすれ", 41), ("その他", 29),
        ]
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "横棒グラフ（項目名が長いときの比較）",
        "項目名が長い、または項目数が多いときに縦棒の代わりに使う。",
        [
            "・項目名が8文字を超えるなら縦棒より横棒。縦棒だと名前が斜めになり読めない",
            "・上から降順に並べる（Excelは既定で逆順になるので軸を反転する）",
            "・項目数は12程度まで。超えるなら上位N＋その他にまとめる",
        ],
        ref, 2,
    )
    ch = BarChart()
    ch.type = "bar"
    ch.gapWidth = 50
    data = Reference(ws, min_col=3, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    solid(ch.series[0], NAVY)
    value_labels(ch)
    ch.legend = None
    ch.y_axis.scaling.orientation = "maxMin"  # 上から降順に見せる
    style_chart(ch, "不良項目別の発生件数", x_title="件数")
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "02_横棒グラフ.xlsx")


def t_line(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "折れ線グラフ"
    months = [f"{m}月" for m in range(1, 13)]
    actual = [312, 298, 305, 287, 271, 264, 258, 249, 233, 221, 208, 196]
    last_year = [330, 325, 318, 322, 310, 305, 301, 298, 292, 288, 281, 277]
    rows = [["月", "今年度", "前年度"]] + [
        [m, a, l] for m, a, l in zip(months, actual, last_year)
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "折れ線グラフ（時系列の推移）",
        "時間に沿った変化と傾向を見る。項目間の比較には使わない。",
        [
            "・横軸は必ず時間。時間でないものを折れ線にしない（棒グラフを使う）",
            "・系列は4本まで。それ以上は読めないので、主役1本＋グレーの背景に分ける",
            "・主張したい系列を濃紺、比較対象（前年など）をグレーにする",
            "・縦軸を0から始めない場合は、その旨を軸か注記に明示する",
        ],
        ref, 3,
    )
    ch = LineChart()
    data = Reference(ws, min_col=3, max_col=4, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    solid(ch.series[0], NAVY, line=True)
    solid(ch.series[1], GRAY, line=True)
    ch.series[1].graphicalProperties.line.dashStyle = "dash"
    for s in ch.series:
        s.smooth = False
        s.marker = Marker(symbol="none")
    style_chart(ch, "不良件数の推移", y_title="件数")
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "03_折れ線グラフ.xlsx")


def t_stacked_bar(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "積み上げ棒グラフ"
    rows = [["工程", "寸法不良", "外観不良", "組付不良"]] + [
        ["第1工程", 42, 28, 12], ["第2工程", 31, 44, 19],
        ["第3工程", 18, 22, 51], ["第4工程", 9, 14, 33],
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "積み上げ棒グラフ（合計と内訳を同時に見る）",
        "合計の大小と、その内訳の構成を1つのグラフで示す。",
        [
            "・内訳は4つまで。増えるほど各層が薄くなり比較できない",
            "・一番下の層は基準がそろうので比較しやすい。主役をここに置く",
            "・合計でなく構成比を比べたいなら 05_帯グラフ を使う",
        ],
        ref, 4,
    )
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.gapWidth = 60
    data = Reference(ws, min_col=3, max_col=5, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s, c in zip(ch.series, [NAVY, BLUE, LIGHT]):
        solid(s, c)
    style_chart(ch, "工程別・不良種別の内訳", y_title="件数")
    chart_font(ch)
    ws.add_chart(ch, "G4")
    return save(wb, outdir, "04_積み上げ棒グラフ.xlsx")


def t_stacked_100(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "帯グラフ"
    rows = [["時期", "寸法不良", "外観不良", "組付不良"]] + [
        ["改善前", 52, 31, 17], ["対策後1ヶ月", 38, 34, 28],
        ["対策後3ヶ月", 21, 36, 43], ["対策後6ヶ月", 14, 33, 53],
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "帯グラフ（構成比の比較・変化）",
        "合計を100%にそろえ、内訳の割合が時期や条件でどう変わるかを見る。",
        [
            "・合計の大きさは見えなくなる。合計も見せたいなら 04_積み上げ棒グラフ",
            "・円グラフを複数並べるより、帯グラフを縦に並べる方が比較しやすい",
            "・並び順は全ての帯で同じにする。順序を変えると比較できない",
        ],
        ref, 4,
    )
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "percentStacked"
    ch.overlap = 100
    ch.gapWidth = 50
    data = Reference(ws, min_col=3, max_col=5, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s, c in zip(ch.series, [NAVY, BLUE, LIGHT]):
        solid(s, c)
    ch.y_axis.scaling.orientation = "maxMin"
    style_chart(ch, "不良種別の構成比の変化", x_title="構成比")
    chart_font(ch)
    ws.add_chart(ch, "G4")
    return save(wb, outdir, "05_帯グラフ.xlsx")


def t_pie(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "円グラフ"
    rows = [["区分", "件数"]] + [
        ["寸法不良", 186], ["外観不良", 152], ["組付不良", 118], ["その他", 70],
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "円グラフ（1時点の構成比）",
        "1つの母数の内訳を直感的に示す。比較には向かない。",
        [
            "・区分は5つまで。それ以上は上位4＋その他にまとめる",
            "・2時点を比べたいなら円を2つ並べず 05_帯グラフ を使う",
            "・12時の位置から時計回りに降順で並べる",
            "・3D・ドーナツの装飾は使わない。面積の比較が狂う",
        ],
        ref, 2,
    )
    ch = PieChart()
    data = Reference(ws, min_col=3, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    value_labels(ch, percent=True, cat_name=True)
    style_chart(ch, "不良区分の構成比", w=16, h=11)
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "06_円グラフ.xlsx")


def t_combo(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "複合グラフ"
    months = [f"{m}月" for m in range(1, 13)]
    volume = [1200, 1180, 1250, 1300, 1280, 1320, 1290, 1350, 1400, 1380, 1420, 1450]
    rate = [4.2, 3.9, 3.6, 3.3, 3.1, 2.8, 2.6, 2.4, 2.1, 1.9, 1.7, 1.5]
    rows = [["月", "生産数", "不良率(%)"]] + [
        [m, v, r] for m, v, r in zip(months, volume, rate)
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "複合グラフ（量と率を重ねる）",
        "単位の違う2つの指標を1枚に重ねる。棒＝量、折れ線＝率が定石。",
        [
            "・第2軸を使うときは、どちらの系列がどちらの軸かを凡例か軸名で明示する",
            "・第2軸の目盛りの取り方で印象が変わる。意図的な誇張をしない",
            "・3つ以上の指標を重ねない。読めなくなる",
        ],
        ref, 3,
    )
    bar = BarChart()
    bar.type = "col"
    bar.gapWidth = 80
    d1 = Reference(ws, min_col=3, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    bar.add_data(d1, titles_from_data=True)
    bar.set_categories(cats)
    solid(bar.series[0], LIGHT)
    bar.y_axis.title = "生産数"

    line = LineChart()
    d2 = Reference(ws, min_col=4, min_row=4, max_row=ref["last_row"])
    line.add_data(d2, titles_from_data=True)
    solid(line.series[0], ACCENT, line=True)
    line.series[0].smooth = False
    line.y_axis.axId = 200
    line.y_axis.title = "不良率(%)"
    line.y_axis.crosses = "max"
    bar += line
    style_chart(bar, "生産数と不良率の推移")
    chart_font(bar)
    ws.add_chart(bar, "F4")
    return save(wb, outdir, "07_複合グラフ.xlsx")


def t_scatter(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "散布図"
    xs = [18.2, 19.1, 20.4, 21.0, 22.3, 23.1, 24.0, 25.2, 26.1, 27.4,
          28.0, 29.3, 30.1, 31.2, 32.0, 33.4, 34.1, 35.0, 36.2, 37.1,
          38.0, 39.3, 40.1, 41.0, 42.2, 43.1, 44.0, 45.3, 46.1, 47.0]
    ys = [0.82, 0.88, 0.91, 0.95, 1.02, 1.05, 1.11, 1.18, 1.21, 1.29,
          1.31, 1.38, 1.42, 1.49, 1.52, 1.61, 1.64, 1.70, 1.78, 1.81,
          1.88, 1.95, 1.99, 2.04, 2.12, 2.18, 2.22, 2.30, 2.34, 2.41]
    rows = [["炉内温度(℃)", "反り量(mm)"]] + [[x, y] for x, y in zip(xs, ys)]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "散布図（2つの量の関係）",
        "2つの量に関係があるかを見る。要因解析で使う。",
        [
            "・最低30組、できれば50組。少ないと見かけの相関が出る",
            "・相関があっても因果とは限らない。第3の要因を疑う",
            "・層別すると相関が消える/現れることがある。条件別に色を分けて確認する",
            "・近似直線は目安。数式と決定係数を出すなら根拠として扱う",
        ],
        ref, 2,
    )
    ch = ScatterChart()
    ch.style = None
    xref = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    yref = Reference(ws, min_col=3, min_row=4, max_row=ref["last_row"])
    s = Series(yref, xref, title_from_data=True)
    s.marker = Marker(symbol="circle", size=6)
    s.marker.graphicalProperties.solidFill = NAVY
    s.marker.graphicalProperties.line.noFill = True
    s.graphicalProperties.line.noFill = True
    s.trendline = Trendline(trendlineType="linear", dispRSqr=True, dispEq=True)
    ch.series.append(s)
    ch.legend = None
    style_chart(ch, "炉内温度と反り量の関係", x_title="炉内温度(℃)", y_title="反り量(mm)")
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "08_散布図.xlsx")


def t_radar(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "レーダーチャート"
    rows = [["評価軸", "自社", "競合A"]] + [
        ["品質", 4, 3], ["コスト", 3, 4], ["納期", 5, 3],
        ["対応速度", 4, 2], ["技術力", 3, 5], ["サポート", 5, 3],
    ]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "レーダーチャート（多項目の特性比較）",
        "5項目以上の評価軸で、全体の形として強み弱みを見る。",
        [
            "・軸は5〜8。少ないと形にならず、多いと読めない",
            "・全軸の尺度と向きをそろえる（大きいほど良い、で統一する）",
            "・比較は2〜3系列まで。面を塗るのは1系列だけにする",
            "・軸の並び順で形が変わる。恣意的に並べ替えない",
        ],
        ref, 3,
    )
    ch = RadarChart()
    ch.type = "marker"
    data = Reference(ws, min_col=3, max_col=4, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    solid(ch.series[0], NAVY, line=True)
    solid(ch.series[1], GRAY, line=True)
    ch.series[1].graphicalProperties.line.dashStyle = "dash"
    style_chart(ch, "自社と競合Aの特性比較", w=16, h=12)
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "09_レーダーチャート.xlsx")


def t_waterfall(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "ウォーターフォール"
    # 積み上げ棒の下段を透明にして増減を表現する（Excel 2016 未満でも開ける）
    steps = [
        ("改善前", 0, 480, "start"),
        ("段取り見直し", 360, 120, "down"),
        ("治具変更", 250, 110, "down"),
        ("条件最適化", 180, 70, "down"),
        ("新規要因", 180, 40, "up"),
        ("改善後", 0, 220, "end"),
    ]
    rows = [["区分", "土台(非表示)", "増減"]] + [[n, b, v] for n, b, v, _ in steps]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "ウォーターフォール（増減の内訳）",
        "始点から終点までの間に、何がどれだけ効いたかを分解して示す。",
        [
            "・「土台(非表示)」列は塗りなしにして見えなくする。値を消さないこと",
            "・増減の合計が終点と一致するか必ず検算する。合わないと図が嘘になる",
            "・減少と増加で色を変える。ここでは減少=濃紺、増加=赤",
            "・要因は6つまで。細かい要因はその他にまとめる",
        ],
        ref, 3,
    )
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.gapWidth = 40
    data = Reference(ws, min_col=3, max_col=4, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.noFill = True   # 土台を透明に
    ch.series[0].graphicalProperties.line.noFill = True
    solid(ch.series[1], NAVY)
    value_labels(ch)
    ch.legend = None
    style_chart(ch, "不良件数の増減内訳", y_title="件数")
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "10_ウォーターフォール.xlsx")


def t_area(outdir):
    wb = Workbook()
    ws = wb.active
    ws.title = "面グラフ"
    months = [f"{m}月" for m in range(1, 13)]
    a = [120, 128, 133, 141, 150, 158, 165, 172, 180, 188, 195, 203]
    b = [80, 82, 85, 88, 90, 93, 95, 98, 100, 103, 105, 108]
    rows = [["月", "A製品", "B製品"]] + [[m, x, y] for m, x, y in zip(months, a, b)]
    ref = write_table(ws, rows)
    style_sheet(
        ws, "面グラフ（累積量の推移）",
        "時系列で、合計の推移とその内訳を同時に見る。",
        [
            "・折れ線で足りるならそちらを使う。面グラフは重なると読みにくい",
            "・系列は3つまで。大きい系列を下に置く",
            "・縦軸は必ず0から始める",
        ],
        ref, 3,
    )
    ch = AreaChart()
    ch.grouping = "stacked"
    data = Reference(ws, min_col=3, max_col=4, min_row=4, max_row=ref["last_row"])
    cats = Reference(ws, min_col=2, min_row=5, max_row=ref["last_row"])
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s, c in zip(ch.series, [NAVY, LIGHT]):
        solid(s, c)
    style_chart(ch, "製品別の出荷数推移", y_title="出荷数")
    chart_font(ch)
    ws.add_chart(ch, "F4")
    return save(wb, outdir, "11_面グラフ.xlsx")


BUILDERS = [
    t_bar_vertical, t_bar_horizontal, t_line, t_stacked_bar, t_stacked_100,
    t_pie, t_combo, t_scatter, t_radar, t_waterfall, t_area,
]


def _assets_templates():
    """図テンプレートの置き場所。**このスキルの中。**

    スキルは `~/.claude/skills/<name>` から参照される。
    **その外に置いたものは、スキルの一部ではない。**
    実際に、リポジトリ直下に置いたところ
    「スキルだけを配ると雛形が0件」「規約の相対パスが空振り」になった。

    `__file__` は登録のリンク越しに呼ばれるので、**resolve() してから辿る。**
    """
    return pathlib.Path(__file__).resolve().parent.parent / "assets" / "templates"


def main():
    import argparse

    ap = argparse.ArgumentParser(description="グラフテンプレートを生成する")
    ap.add_argument("outdir", nargs="?", default=None, help="出力先")
    ap.add_argument("--latin", default=None,
                    help="欧文フォント名。省くと日本語と同じものを使う")
    ap.add_argument("--font", default=FONT,
                    help="日本語フォント名。**資料テンプレートに合わせる。**"
                         "read_template_style.py で調べた値を渡す")
    args = ap.parse_args()

    # フォントは本スキルで固定しない。資料テンプレートが決める
    globals()["FONT"] = args.font
    globals()["LATIN"] = args.latin

    outdir = Path(args.outdir) if args.outdir else (
        _assets_templates() / "charts"
    )
    outdir.mkdir(parents=True, exist_ok=True)
    print(f"グラフテンプレートを生成: {outdir}  フォント: 日本語={FONT} / 欧文={LATIN or FONT}")
    for fn in BUILDERS:
        fn(outdir)
    print(f"完了: {len(BUILDERS)} 件")


if __name__ == "__main__":
    main()

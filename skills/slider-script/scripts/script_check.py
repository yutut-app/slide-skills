#!/usr/bin/env python3
"""台本を機械で検査する。**書いた後、報告の前に必ず通す。**

    python3 scripts/script_check.py <案件>/script.md --html <案件>/deck.html --minutes 20

## 何を見るか

| 見るもの | 落ちる条件 |
|---|---|
| 枚との対応 | 台本にある枚がスライドに無い／スライドにある枚が台本に無い |
| 尺 | 持ち時間に対する**超過**。**不足も出す**（早く終わるのも失敗） |
| 欄 | 伝えること・つなぎ・本論・見てほしい場所・次へ のどれかが空 |
| 1枚1メッセージ | **「伝えること」が1文でない**（2つ以上を1枚で言おうとしている） |
| 台本にしか無い数字 | **スライドにも数値台帳にも無い数値が台本に出ている** |
| 話さない言い回し | 作り手の工夫・自己評価・作り手の作法・権限を超える言い方・誇張・立場の語 |
| 定義前の用語 | 用語台帳があるとき、定義より前に出てくる語（**要確認。落とさない**） |

**指摘が1件でもあれば異常終了する**（終了コード 1）。直してから報告する。
ただし「要確認」だけのときは落とさない。**機械が判定しきれないものを不合格にすると、
正しい台本を機械に合わせて曲げることになる。**

## 判定しないもの

**「伝えることが全部入っているか」は字数では分からない。**尺が通っても中身は薄くなりうる。
`references/30_check.md` の、伝えることと本論を目で突き合わせる手順を必ず行う。

**話さない枚**（配布のみ・付録）は、本文に `**話さない**` と書く。
欄の必須と尺の集計から外れる。**空欄のまま置かない。**

**話す速さは1分あたり 300字で見積もる**（`references/20_timing.md`）。
`--chars-per-minute` で変えられるが、**速い側に振らない。**

HTML を渡さないと、枚との対応と数字の照合はできない。**その旨を前提欄に出す。**
"""

import argparse
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked

FIELDS = ["伝えること", "つなぎ", "本論", "見てほしい場所", "次へ"]
# 声に出さない欄。書いてよいが、尺にも本文の検査にも入れない
SILENT_FIELDS = ["想定質問", "メモ"]
SKIP_MARK = "**話さない**"

# **単位付き、または2桁以上の数だけを見る。**「上位2項目」の「2」まで拾うと、
# 雑音で本物が埋もれる（作ったときに実際に誤検出した）
_NUM = re.compile(r"(?<![\d.])(\d+(?:[.,]\d+)?)\s*"
                  r"(%|％|パーセント|件|円|人|分|秒|時間|日|台|回|倍|万|億|"
                  r"メートル|ミリ|キロ|グラム|kg|g|mm|cm|m|min|h)?")
# 「1分あたり」「毎分」のような割合の言い方は、主張している数値ではない
_RATE_AFTER = ("あたり", "当たり", "ごと", "間隔", "おき")
_RATE_BEFORE = ("毎", "／", "/")


def _digits(s: str) -> str:
    return unicodedata.normalize("NFKC", s)


def norm_num(value: str) -> str:
    """比較用に値だけを残す。**単位と表記の違いで落とさない。**

    スライドが「3.2％」、台本が「3.2パーセント」、あるいは「23.4」と「23.4件」を
    別物と判定すると、正しい台本を数字に合わせて曲げることになる（実際に起きた）。
    """
    v = _digits(value)
    if re.fullmatch(r"\d{1,3}(,\d{3})+", v):     # 桁区切り
        v = v.replace(",", "")
    else:                                         # 区切りのカンマは切る
        v = v.split(",")[0]
    try:
        f = float(v)
    except ValueError:
        return v
    return f"{f:g}"


def numbers_in(text: str):
    """{正規化した値} を返す。単位は落とし、割合の言い方は数えない。"""
    t = _digits(text)
    out = set()
    for m in _NUM.finditer(t):
        after = t[m.end():m.end() + 3]
        before = t[max(0, m.start() - 1):m.start()]
        if any(after.startswith(a) for a in _RATE_AFTER):
            continue
        if before in _RATE_BEFORE:
            continue
        raw, unit = m.group(1), m.group(2)
        val = norm_num(raw)
        # 単位も2桁も無い一桁の数（順位・個数）は雑音
        if not unit and len(raw.replace(".", "").replace(",", "")) < 2:
            continue
        out.add(val)
    return out


def rows_of(md: Path):
    """Markdown の表を [{列: 値}] にする。"""
    rows, head = [], None
    for line in md.read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if set("".join(cells)) <= set("-: "):
            continue
        if head is None:
            head = cells
            continue
        rows.append(dict(zip(head, cells)))
    return rows


def terms_of(md: Path):
    """用語台帳から [(使う語, 初出で言う定義)]。定義の列は無くてよい。"""
    out = []
    for d in rows_of(md):
        w = d.get("使う語") or d.get("語") or ""
        if w:
            out.append((w, d.get("定義", "")))
    return out


def numbers_of(md: Path):
    """数値台帳から {正規化した値}。**スライドの別の枚にある値を誤検出しないため。**"""
    out = set()
    for d in rows_of(md):
        for v in d.values():
            out |= numbers_in(v)
    return out


def phrases_of(md: Path):
    """言い回しの台帳から [(区分, 語)]。"""
    return [(d.get("区分", "—"), d.get("使わない語") or d.get("語"))
            for d in rows_of(md) if d.get("使わない語") or d.get("語")]


def slides_of_script(md: Path):
    """[(枚番号, 見出し, 本文)]。`## <n>. <見出し>` を1枚とする。"""
    out, cur, buf = [], None, []
    for line in md.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^##\s+(\d+)\.?\s*(.*)$", line)
        if m:
            if cur:
                out.append((cur[0], cur[1], "\n".join(buf)))
            cur, buf = (int(m.group(1)), m.group(2).strip()), []
        elif cur:
            buf.append(line)
    if cur:
        out.append((cur[0], cur[1], "\n".join(buf)))
    return out


def slides_of_html(html: Path):
    """[(枚番号, 見える文字)]。**UTF-8 で明示的に読む**（charset 無しの誤読を防ぐ）。"""
    import lxml.html as LH

    doc = LH.document_fromstring(html.read_text(encoding="utf-8"))
    for bad in doc.xpath("//script|//style"):
        bad.getparent().remove(bad)
    sls = doc.xpath(
        "//*[contains(concat(' ', normalize-space(@class), ' '), ' slide ')]") or [doc]
    return [(i, re.sub(r"\s+", " ", " ".join(s.itertext()).replace(" ", " ")))
            for i, s in enumerate(sls, 1)]


def _spoken_lines(body: str):
    """声に出す行だけ。欄の見出し・注記・声に出さない欄は外す。"""
    for line in body.splitlines():
        line = line.strip()
        if not line or line.startswith("<!--"):
            continue
        if any(line.startswith(f"**{f}**") for f in SILENT_FIELDS):
            continue
        yield re.sub(r"^\*\*(%s)\*\*[：:]\s*" % "|".join(FIELDS), "", line)


def spoken_text(body: str) -> str:
    return " ".join(_spoken_lines(body))


def spoken_chars(body: str) -> int:
    return sum(len(re.sub(r"[#*`\-|]", "", l)) for l in _spoken_lines(body))


def field_value(body: str, name: str) -> str:
    m = re.search(r"^\*\*%s\*\*[：:]\s*(.*)$" % name, body, re.M)
    return m.group(1).strip() if m else ""


def main():
    here = Path(__file__).resolve().parent.parent
    ap = argparse.ArgumentParser(
        description="台本を機械で検査する。書いた後、報告の前に必ず通す")
    ap.add_argument("script", help="台本（script.md）")
    ap.add_argument("--html", help="スライドの HTML。枚との対応と数字の照合に使う")
    ap.add_argument("--minutes", type=float, help="持ち時間（分）。質疑は含めない")
    ap.add_argument("--chars-per-minute", type=float, default=300.0,
                    help="話す速さ（既定 300字/分）。**速い側に振らない**")
    ap.add_argument("--terms", help="用語台帳（terms.md）。定義前に出る語を見る")
    ap.add_argument("--numbers",
                    help="数値台帳。**スライドに無いが正しい値**を誤検出から外す")
    ap.add_argument("--phrases", default=str(here / "assets" / "phrases.md"),
                    help="話さない言い回しの台帳。案件固有の語はコピーして差し替える")
    args = ap.parse_args()

    script = slides_of_script(Path(args.script))
    if not script:
        print("**台本に枚が見つからない。**`## 3. 見出し` の形で枚を区切る")
        return checked.summary("script_check", 0, "枚", 1, {"台本": args.script})

    findings, notes = [], []        # notes は「要確認」。終了コードを落とさない
    spoken = [(n, h, b) for n, h, b in script if SKIP_MARK not in b]
    silent = len(script) - len(spoken)

    # 欄の空きと、1枚1メッセージ
    for no, _, body in spoken:
        missing = [f for f in FIELDS
                   if f"**{f}**" not in body or not field_value(body, f)]
        if no == 1:
            missing = [f for f in missing if f != "つなぎ"]
        if missing:
            findings.append((no, "欄が空", "／".join(missing)))
        msg = field_value(body, "伝えること")
        if msg and (len(re.findall(r"[。．]", msg)) > 1 or len(msg) > 60):
            findings.append((no, "伝えることが1つでない",
                             "**1枚で言うのは1つ。**分けるか、枚を分ける"))

    # 尺（話さない枚は数えない）
    total = sum(spoken_chars(b) for _, _, b in spoken)
    minutes = args.minutes
    if minutes:
        budget = minutes * args.chars_per_minute
        ratio = total / budget if budget else 0
        if ratio > 1.1:
            findings.append(("—", "尺の超過",
                             f"{total:.0f}字 / 目安 {budget:.0f}字（{ratio:.2f} 倍）。**削る**"))
        elif ratio < 0.8:
            findings.append(("—", "尺の不足",
                             f"{total:.0f}字 / 目安 {budget:.0f}字（{ratio:.2f} 倍）。**足す**"))

    # 話さない言い回し
    ph_path = Path(args.phrases)
    phrases = phrases_of(ph_path) if ph_path.exists() else []
    for no, _, body in spoken:
        t = spoken_text(body)
        for kind, word in phrases:
            if word in t:
                findings.append((no, f"話さない言い回し（{kind}）",
                                 f"「{word}」。**事実と判断だけを話す**"))

    # 枚との対応・数字
    allowed = numbers_of(Path(args.numbers)) if args.numbers else set()
    html_slides = slides_of_html(Path(args.html)) if args.html else []
    if html_slides:
        s_no = {n for n, _, _ in script}
        h_no = {n for n, _ in html_slides}
        for n in sorted(s_no - h_no):
            findings.append((n, "枚がスライドに無い", "台本にあるが、スライドに対応する枚が無い"))
        for n in sorted(h_no - s_no):
            findings.append((n, "台本に無い枚", "スライドにあるが、台本が書かれていない"))
        h_text = {n: t for n, t in html_slides}
        for n, _, body in spoken:
            if n not in h_text:
                continue
            extra = numbers_in(spoken_text(body)) - numbers_in(h_text[n]) - allowed
            for x in sorted(extra):
                findings.append((n, "台本にしか無い数字",
                                 f"「{x}」がスライドに無い。**根拠を聞かれて答えられない**"))

    # 定義前の用語（要確認。機械では言い切れない）
    if args.terms:
        seen = set()
        for n, _, body in spoken:
            t = spoken_text(body)
            for w, definition in terms_of(Path(args.terms)):
                if w not in t or w in seen:
                    continue
                seen.add(w)
                core = re.sub(r"\s+", "", definition)[:8]
                if core and core in re.sub(r"\s+", "", t):
                    continue
                notes.append((n, "定義前の用語（要確認）",
                              f"「{w}」の初出。**その場で定義しているか目で見る**"))

    print(f"検査: {Path(args.script).name}  台本 {len(script)} 枚"
          f"（話す {len(spoken)} 枚 / 話さない {silent} 枚）  話す文字数 {total:.0f}字")
    if findings or notes:
        print("\n| 枚 | 種類 | 中身 |")
        print("|---|---|---|")
        for f in findings + notes:
            print(f"| {f[0]} | {f[1]} | {f[2]} |")
    if not findings:
        print("\n**機械の指摘なし。**この後、`references/30_check.md` に従い、"
              "**伝えることと本論を突き合わせ**、聞き手の目で通して読む。")

    code = checked.summary("script_check", len(script), "枚", len(findings), {
        "持ち時間": f"{minutes:.0f}分" if minutes else "**渡されていない**",
        "速さ": f"{args.chars_per_minute:.0f}字/分",
        "スライド": Path(args.html).name if args.html else "**渡されていない（枚と数字は見ていない）**",
        "数値台帳": Path(args.numbers).name if args.numbers else "無し",
        "言い回しの台帳": f"{len(phrases)} 語",
        "要確認": f"{len(notes)} 件",
    })
    # **指摘があれば落とす。**サマリ行は「見た」ことしか示さない。
    # 「要確認」では落とさない。機械が判定しきれないものを不合格にすると、
    # 正しい台本を機械に合わせて曲げることになる
    return code or (1 if findings else 0)


if __name__ == "__main__":
    sys.exit(main())

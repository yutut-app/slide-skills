#!/usr/bin/env python3
"""日本語の行の崩れと、画像の劣化を検出する。**画像で1枚ずつ見る前に通す。**

    python3 scripts/jp_wrap_check.py <検査する>.pptx [--before <元の>.pptx]

| 見るもの | 区分 | 理由 |
|---|---|---|
| 行頭に読点・句点・閉じかっこが来る | **要確認** | 折り返しは描画で決まるので、推定は外れうる |
| 最後の行が1文字だけ（「る」だけが落ちる等） | **要確認** | 同上 |
| 見出しに改行が残っている | **要確認** | 文を書き換えたとき、元の改行が残って1行目が押し出された |
| 画像の縦横比が、元の画像（切り取り後）とずれる | **指摘** | 数値で決まる。**1% を超えたら**落とす（実績 3% で崩れて見えた） |
| `--before` より画像の画素数が減った | **指摘** | 画像を作り直したときに解像度が下がった |
| 表示の大きさに対して解像度が低い（150dpi 未満） | 要確認 | 投影では荒れて見えることがある |

**要確認は落とさない。**確定は `qa_render.py` などで画像にして目で見る。
指摘があれば異常終了する（終了コード1）。
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
from fit_check import TEXT_MARGIN_INCH, body_pt, char_width_units

HEAD_NG = set("、。，．,.）)」』】〕〉》］}・ー！？!?")   # 行頭に来てはいけない
EMU_PER_INCH = 914400
ASPECT_TOL = 0.01
DPI_LOW = 150


def wrap(para: str, cols_units: float):
    """全角1・半角0.5で、**左から詰めて**折り返す（PowerPoint の禁則は考えない推定）。"""
    lines, cur, w = [], "", 0.0
    for ch in para:
        cw = char_width_units(ch)
        if w + cw > cols_units and cur:
            lines.append(cur)
            cur, w = "", 0.0
        cur += ch
        w += cw
    lines.append(cur)
    return lines


def text_checks(sh, no, notes):
    if not sh.has_text_frame or not sh.width:
        return
    tf = sh.text_frame
    pt = body_pt(sh)
    left = (tf.margin_left if tf.margin_left is not None else int(TEXT_MARGIN_INCH * EMU_PER_INCH))
    right = (tf.margin_right if tf.margin_right is not None else int(TEXT_MARGIN_INCH * EMU_PER_INCH))
    usable_in = (sh.width - left - right) / EMU_PER_INCH
    if usable_in <= 0 or pt <= 0:
        return
    cols = usable_in / (pt / 72)
    is_title = sh.is_placeholder and "TITLE" in str(sh.placeholder_format.type)
    name = (sh.name or "")[:14]
    for para in tf.paragraphs:
        raw = para.text            # runs を連結すると改行（a:br）が落ちる。実績: 見出しの改行を見逃した
        if is_title and ("\v" in raw or "\x0b" in raw):
            notes.append((no, name, "見出しに改行が残っている",
                          "文を書き換えたとき、元の改行で1行目が押し出されていないか"))
        for seg in raw.replace("\x0b", "\v").split("\v"):
            seg = seg.strip()
            if not seg:
                continue
            lines = wrap(seg, cols)
            if len(lines) < 2:
                continue
            for i, ln in enumerate(lines[1:], 2):
                if ln and ln[0] in HEAD_NG:
                    notes.append((no, name, "行頭に読点・かっこ（推定）",
                                  f"{i}行目が「{ln[0]}」で始まる：…{lines[i-2][-6:]}／{ln[:6]}…"))
                    break
            last = lines[-1]
            if char_width_units(last) <= 1.0:
                notes.append((no, name, "1文字だけの折り返し（推定）",
                              f"最後の行が「{last}」だけ：…{lines[-2][-8:]}／{last}"))
    if is_title and len(tf.paragraphs) > 1 and sum(1 for p in tf.paragraphs if p.text.strip()) > 1:
        notes.append((no, name, "見出しが複数の段落", "元の改行が段落として残っていないか"))


def pictures(prs):
    out = []
    for no, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if sh.shape_type == 13:          # PICTURE
                out.append((no, sh))
    return out


def pic_info(sh):
    pw, ph = sh.image.size
    cl, cr, ct, cb = sh.crop_left, sh.crop_right, sh.crop_top, sh.crop_bottom
    vis_w, vis_h = pw * (1 - cl - cr), ph * (1 - ct - cb)
    return pw, ph, vis_w, vis_h


def main():
    ap = argparse.ArgumentParser(description="日本語の行の崩れと画像の劣化を検出する")
    ap.add_argument("pptx")
    ap.add_argument("--before", help="元の pptx。画像の解像度が下がっていないかを比べる")
    args = ap.parse_args()
    from pptx import Presentation

    prs = Presentation(Path(args.pptx).expanduser().resolve())
    findings, notes, n = [], [], 0
    for no, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            n += 1
            text_checks(sh, no, notes)

    pics = pictures(prs)
    for no, sh in pics:
        pw, ph, vw, vh = pic_info(sh)
        if not (sh.width and sh.height and vw and vh):
            continue
        native = vw / vh
        shown = sh.width / sh.height
        diff = abs(shown / native - 1)
        if diff > ASPECT_TOL:
            findings.append((no, sh.name[:14], "画像の縦横比がずれる",
                             f"{diff * 100:.1f}%（許容 {ASPECT_TOL * 100:.0f}%）。切り取りか枠の比を合わせる"))
        dpi = vw / (sh.width / EMU_PER_INCH)
        if dpi < DPI_LOW:
            notes.append((no, sh.name[:14], "画像の解像度が低い",
                          f"約 {dpi:.0f}dpi（目安 {DPI_LOW}）。投影で荒れないか"))

    if args.before:
        before = pictures(Presentation(Path(args.before).expanduser().resolve()))
        by_slide = {}
        for no, sh in before:
            by_slide.setdefault(no, []).append(sh)
        for no, sh in pics:
            olds = by_slide.get(no, [])
            if not olds:
                continue
            old = min(olds, key=lambda o: abs(o.left - sh.left) + abs(o.top - sh.top))
            _, _, ovw, ovh = pic_info(old)
            _, _, vw, vh = pic_info(sh)
            if vw * vh < ovw * ovh * 0.95:
                findings.append((no, sh.name[:14], "元より解像度が下がった",
                                 f"{int(ovw)}×{int(ovh)} → {int(vw)}×{int(vh)} 画素（見えている範囲）"))

    if findings or notes:
        print("| 枚 | 図形 | 種類 | 中身 |")
        print("|---|---|---|---|")
        for f in findings + notes:
            print(f"| {f[0]} | {f[1]} | {f[2]} | {f[3]} |")
    else:
        print("**指摘なし。**この後、画像にして1枚ずつ目で見る（推定は外れうる）。")

    code = checked.summary("jp_wrap_check", n, "図形", len(findings), {
        "画像": f"{len(pics)} 件",
        "比較": Path(args.before).name if args.before else "**元の pptx を渡していない（解像度の低下は見ていない）**",
        "要確認": f"{len(notes)} 件（折り返しは推定）",
    })
    return code or (1 if findings else 0)


if __name__ == "__main__":
    sys.exit(main())

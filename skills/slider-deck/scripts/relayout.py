#!/usr/bin/env python3
"""枠の位置を、HTML の座標を写さずに組み直す。

**なぜ要るか。** HTML は px の絶対座標で書かれているが、PowerPoint は
枠の中の文字を**指定した高さより縦に伸ばす**（書体の置換・表の行高・
既定文字サイズへの落ち）。写した座標は、伸びたぶんだけ下の枠と重なる。
実績では18枚中12枚が崩れ、そのすべてが**縦の伸び**で、横位置は合っていた。

**やること。** 座標を「位置」ではなく**並び順と行のまとまり**として読み、
中身の枠を**行に束ね、余白を揃えて敷き直す。**重なりは作りようがなくなる。

**動かさないもの（固定）。**
  - 画像を含む枠（ロゴ）
  - 上端・下端に接した薄い帯（見出し帯・ページ番号）
  - `data-pin` が付いた枠
器が崩れると、直す手間のほうが大きい。**得が無いので動かさない。**

使う側は `build_slide(..., relayout=True)`。座標の辞書を書き換えるだけで、
書体・表・図の作り方には触らない。
"""

# 帯とみなす高さ（スライド高に対する比）と、端に接しているとみなす距離
BAND_RATIO = 0.10
EDGE_PX = 8.0
# 余白。**行の間と枠の間で別に持つ。**同じにすると行の切れ目が見えなくなる
GUTTER_X = (8.0, 40.0)
GUTTER_Y = (10.0, 40.0)
SIDE_MARGIN = 0.04       # 左右の余白がまったく無いときに使う比


def _f(style, key, default=None):
    import html_abs as HA
    return HA.px(style.get(key), default)


def _box(style):
    """(left, top, width, height)。取れないものは None。"""
    return (_f(style, "left"), _f(style, "top"),
            _f(style, "width"), _f(style, "height"))


def is_pinned(el, style, slide_w, slide_h):
    """動かさない枠か。**器（ロゴ・帯）は動かしても得が無い。**"""
    if el.get("data-pin") is not None:
        return True
    if el.tag == "img" or el.xpath(".//img"):
        return True
    left, top, w, h = _box(style)
    if top is None or h is None:
        return True                       # 寸法が読めないものは触らない
    if h <= slide_h * BAND_RATIO and (top <= EDGE_PX
                                      or top + h >= slide_h - EDGE_PX):
        return True
    return False


def _rows(items):
    """縦に重なるものを1行に束ねる。**重なりの深さで見る。**

    上端だけで切ると、高さの違う隣どうしが別の行に落ちる。
    """
    rows = []
    for it in sorted(items, key=lambda t: (_f(t[1], "top", 0.0),
                                           _f(t[1], "left", 0.0))):
        top, h = _f(it[1], "top", 0.0), _f(it[1], "height", 0.0) or 0.0
        placed = False
        for row in rows:
            rtop = min(_f(x[1], "top", 0.0) for x in row)
            rbot = max(_f(x[1], "top", 0.0) + (_f(x[1], "height", 0.0) or 0.0)
                       for x in row)
            overlap = min(rbot, top + h) - max(rtop, top)
            if overlap > 0 and overlap >= 0.5 * min(h or 1.0, rbot - rtop or 1.0):
                row.append(it)
                placed = True
                break
        if not placed:
            rows.append([it])
    for row in rows:
        row.sort(key=lambda t: _f(t[1], "left", 0.0))
    return rows


def _clamp(v, lo_hi):
    lo, hi = lo_hi
    return max(lo, min(hi, v))


def _median(vals, default):
    vals = sorted(v for v in vals if v is not None)
    if not vals:
        return default
    return vals[len(vals) // 2]


def relayout(items, slide_w, slide_h, warn=None):
    """items = [(el, style), ...]（`.slide` の直下）。style を書き換える。

    戻り値は (動かした数, 固定した数)。
    **入りきらないときは縮めない。**警告に出して、そのまま置く。
    縮めると字が読めなくなり、直すべき箇所が見えなくなる。
    """
    warn = warn if warn is not None else []
    pinned, flow = [], []
    for el, st in items:
        (pinned if is_pinned(el, st, slide_w, slide_h) else flow).append((el, st))
    if not flow:
        return 0, len(pinned)

    # 中身を置ける帯。**固定した枠を避ける。**
    top_edge, bot_edge = 0.0, slide_h
    for _, st in pinned:
        _, t, _, h = _box(st)
        if t is None or h is None:
            continue
        mid = t + h / 2
        if mid < slide_h / 2:
            top_edge = max(top_edge, t + h)
        else:
            bot_edge = min(bot_edge, t)

    lefts = [_f(st, "left") for _, st in flow]
    rights = [(_f(st, "left") or 0.0) + (_f(st, "width") or 0.0) for _, st in flow]
    x0 = min([v for v in lefts if v is not None], default=slide_w * SIDE_MARGIN)
    x1 = max(rights, default=slide_w * (1 - SIDE_MARGIN))
    if x1 <= x0:
        x0, x1 = slide_w * SIDE_MARGIN, slide_w * (1 - SIDE_MARGIN)
    avail_w = x1 - x0

    rows = _rows(flow)

    # 縦の余白は、元の行間の中央値を使う。**作り手が置いた間隔の感じを残す。**
    gaps = []
    for a, b in zip(rows, rows[1:]):
        abot = max((_f(s, "top", 0.0) + (_f(s, "height", 0.0) or 0.0)) for _, s in a)
        btop = min(_f(s, "top", 0.0) for _, s in b)
        gaps.append(btop - abot)
    gy = _clamp(_median(gaps, 16.0), GUTTER_Y)

    heights = [max((_f(s, "height", 0.0) or 0.0) for _, s in row) for row in rows]

    def _avail(g):
        # **上下の余白も引く。**引き忘れると、最後の行がスライドの外へ出る
        return bot_edge - top_edge - g * (len(rows) - 1) - 2 * g

    avail_h = _avail(gy)
    need_h = sum(heights)
    if need_h <= 0:
        return 0, len(pinned)
    if need_h > avail_h:
        # **縮めない。**まず余白を詰め、それでも入らなければ報告する
        gy = _clamp(GUTTER_Y[0], GUTTER_Y)
        avail_h = _avail(gy)
    scale = avail_h / need_h if need_h > avail_h else 1.0
    if scale < 1.0:
        warn.append(f"整列: 中身が縦に {need_h - avail_h:.0f}px 入りきらない。"
                    "**文字量を減らすか、枚を分ける。**枠は縮めていない")
        scale = 1.0

    # 余りは行に配る。**上に寄せて下を空けない。**空きが偏ると寂しく見える
    extra = (avail_h - need_h) / len(rows) if avail_h > need_h else 0.0

    y = top_edge + gy
    moved = 0
    for row, h in zip(rows, heights):
        rh = h * scale + extra
        widths = [(_f(s, "width", 0.0) or 0.0) for _, s in row]
        gx = _clamp(_median(
            [(_f(b[1], "left", 0.0) - (_f(a[1], "left", 0.0)
                                       + (_f(a[1], "width", 0.0) or 0.0)))
             for a, b in zip(row, row[1:])], 16.0), GUTTER_X)
        total_w = sum(widths)
        room = avail_w - gx * (len(row) - 1)
        k = room / total_w if total_w > 0 else 1.0
        x = x0
        for (el, st), w in zip(row, widths):
            st["left"] = f"{x:.1f}px"
            st["top"] = f"{y:.1f}px"
            st["width"] = f"{max(1.0, w * k):.1f}px"
            st["height"] = f"{max(1.0, rh):.1f}px"
            x += w * k + gx
            moved += 1
        y += rh + gy

    return moved, len(pinned)

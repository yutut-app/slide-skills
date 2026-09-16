#!/usr/bin/env python3
"""検査の道具が、**何を・いくつ・どんな前提で見たか**を1行で出す。

## なぜ要るか

**「指摘ゼロ」と「何も見ていない」が、出力で区別できなかった。**
同じ原因が9回起きている（`docs/feedback.md` の原因 C）。

| 実際に起きたこと | 出力 |
|---|---|
| .potx を渡して python-pptx が開けず、検査が走らなかった | 何も出ず「指摘ゼロ」と報告 |
| 差分ツールが行間を一度も比較していなかった | 「差は無い」 |
| 書体が効いていないのに、属性は入っているので通った | 合格 |
| 換算係数が違い、全項目ずれた差分を出した | **自信のある誤答** |

**個々の道具を直しても次の変種で再発する。** 出力の型で塞ぐ。

## 型

```
検査: <道具>  対象 <n><単位>  指摘 <n>件  前提: <key>=<値>, ...
```

- **対象が 0 なら異常終了する**（終了コード 2）。何も見ていないことを成功にしない
- **前提を書く。** 何を仮定して数えたかが分かると、外れたときに気づける
- **報告にはこの行をそのまま貼る**（`<slider-deck>/references/60_qa.md`）。「確認しました」の代わりにならない
"""

import sys

EXIT_NOTHING_CHECKED = 2


def summary(tool: str, checked: int, unit: str, findings: int,
            assumptions: dict = None, out=None) -> int:
    """検査の1行サマリを出す。対象が0なら異常終了コードを返す。

    戻り値をそのまま `sys.exit()` に渡す。
    """
    out = out or sys.stdout
    parts = [f"検査: {tool}", f"対象 {checked}{unit}", f"指摘 {findings}件"]
    if assumptions:
        parts.append("前提: " + ", ".join(f"{k}={v}" for k, v in assumptions.items()))
    print("\n" + "  ".join(parts), file=out)

    if checked <= 0:
        print("**対象が0。何も見ていない。**「問題なし」と報告しない。", file=out)
        print("  入力の形式・パス・拡張子を確かめる"
              "（python-pptx は .potx を開けない、など）", file=out)
        return EXIT_NOTHING_CHECKED
    return 0


def font_state(*names) -> str:
    """書体が手元にあるかを前提欄に書くための文字列。

    **無いまま検査すると、折り返し位置は当てにならない。**
    「確認した」と言えるかどうかがここで決まる。
    """
    import subprocess
    try:
        listed = subprocess.run(["fc-list"], capture_output=True, text=True,
                                timeout=10).stdout.lower()
    except (OSError, subprocess.SubprocessError):
        return "不明"
    miss = [n for n in names if n.lower() not in listed]
    return "そろい" if not miss else "不足(" + "/".join(miss) + ")"

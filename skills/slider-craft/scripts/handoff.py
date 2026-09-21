#!/usr/bin/env python3
"""引き継ぎメモの雛形を作る。版と指紋を機械で埋める。

    python3 scripts/handoff.py <テンプレ>ディレクトリ -o work/HANDOFF.md

**「同じものを用意した」だけでは足りない。** 名前が同じでも中身が違うことがある。
テンプレートの指紋とスキルの版を入れて、受け取った側が突き合わせられるようにする
（`references/49_handoff.md`）。

**中身は人が書く。** ここで埋まるのは機械で決まる欄だけ。
「今どこ」「直したこと」「未決」「触っていないもの」は会話にしか無い。
"""

import argparse
import hashlib
import subprocess
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked

HERE = Path(__file__).resolve().parent


def fingerprint(tpl: Path) -> str:
    """テンプレートの中身の指紋。html・manifest・images を名前順に連結して取る。

    **画像も見る。** 見た目はロゴなどの画像でも決まるので、
    html だけを見ていると**ロゴが違うのに指紋は一致する。**
    実際に、片方が仮のロゴのままでも気づけない状態だった。

    **名前も混ぜる。** 中身だけだと、名前を変えただけの入れ替えを拾えない。
    """
    files = (sorted(tpl.glob("*.html"))
             + sorted(tpl.glob("template-manifest.md"))
             + sorted(p for p in tpl.glob("images/*") if p.is_file()))
    if not files:
        return "**取れない（html が無い）**"
    h = hashlib.sha256()
    for f in files:
        h.update(f.relative_to(tpl).as_posix().encode("utf-8"))
        h.update(b"\0")
        h.update(f.read_bytes())
    return h.hexdigest()[:12]


def source_fingerprint(files) -> str:
    """正本（資料の HTML 一式）の指紋。**引き継ぎ一式が古くなったかを見るため。**

    正本を直すたびに引き継ぎ一式は陳腐化するが、作り直すきっかけが無い。
    実績: 引き継ぎが第3版のまま、正本だけが先へ進んでいた。
    **メモに正本の指紋を書き、今の正本と比べれば機械的に分かる。**
    """
    h = hashlib.sha256()
    for f in sorted(Path(x) for x in files):
        h.update(f.name.encode("utf-8"))
        h.update(b"\0")
        h.update(f.read_bytes())
    return h.hexdigest()[:12]


def skill_version() -> str:
    """スキルの版。

    **git があるとは限らない。** コピーで配ると履歴が付いてこないので、
    `git log` は空を返す。実際に「取れない」と出て、
    **引き継ぎメモの版の突き合わせができなかった。**

    そこで**配布物に `VERSION` を入れておき、無ければそれを読む。**
    """
    try:
        r = subprocess.run(["git", "-C", str(HERE), "log", "-1", "--format=%h"],
                           capture_output=True, text=True, timeout=10)
        if r.stdout.strip():
            return r.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        pass
    # スキル直下 → その上 と順に見る。**スキルだけ配られることがある**
    for base in (HERE.parent, HERE.parent.parent, HERE.parent.parent.parent):
        f = base / "VERSION"
        if f.exists():
            return f.read_text(encoding="utf-8").strip().splitlines()[0]
    return "**取れない**"


TEMPLATE = """# 引き継ぎ

版: rev<n>（{today}）
渡す人 → 受け取る人: <誰> → <誰>
**スキルの版: {version}**  ← 違えば挙動が違う。指摘が再現しないときはここを疑う

**正本の指紋: {src_fp}**  ← 正本を直したら、この引き継ぎ一式は古い（`--stale` で確かめる）

---
# 現在の状態（書き換える）

## このメモの役割
**相手に渡すメモ**（受け取った人が続きから入るためのもの）。
自分用の続きメモとは別物。**食い違ったら、こちら（渡した時点の事実）が正。**

## 今どこ
段階: <0 / 1 / 2>
<段階1の終了宣言が済んでいるかを書く>

## テンプレート
名前: {name}
スライドサイズ: {size}
**指紋: {fp}**  ← 受け取った側は自分で計算して突き合わせる
同梱した / 相手が持っている（どちらかを書く）

## 正本を作る手順（再現の手順）
**続きから入る人が、同じ正本を作り直せるように書く。**順番と、各段で何が変わるか。
**正本を上書きするスクリプトは、必ずその旨を書く。**黙って上書きされると、
直した内容が消えても気づけない。
| 順 | コマンド | 何が変わるか | 正本を上書きするか |
|---|---|---|---|
| 1 | <コマンド> | <出力> | する / しない |

## 直近で直したこと
| スライド | 直したこと |
|---|---|
| <n> | <内容> |

## 図
<まだ入れていない / rev<n> の pptx に入れた>
**図を入れた pptx は再変換すると消える。**入れた版を必ず書く。

## スキルに当てた直し
**このスキル自身の不具合を直したなら、ここに書いて配布元へ報告する。**
書かないと、**次の配布で黙って消える**（実際に2回消えた）。
- <直した場所と、何が起きていたか>

## 未決
- <残っている判断・待っているデータ>

## 積み残し
**「直さない」と書くときは、根拠を添える。**
「仕様なので直さない」のか「**まだ原因を調べていない**」のかを区別する。
調べていないなら、そう書く。**仕様と誤認して固定された例がある。**

## 触っていないもの
スライド <番号を列挙>

## 同梱した台帳
**用語台帳（terms.md）と数値台帳（numbers.md）も渡す。**受け取った側が同じ語・同じ数字で直せる。
- terms.md / numbers.md（無ければ、渡す前に作る）

---
# 履歴（追記する。消さない）
続きから入る人は、上の「現在の状態」だけ読めばよい。
| 日付 | 版 | したこと |
|---|---|---|
| {today} | rev<n> | <内容> |
"""


def main():
    ap = argparse.ArgumentParser(description="引き継ぎメモの雛形を作る。始める前の版の突き合わせと、返すときのメモ")
    ap.add_argument("template", help="テンプレートのディレクトリ")
    ap.add_argument("-o", "--out", help="書き出し先。省くと標準出力")
    ap.add_argument("--fingerprint", action="store_true",
                    help="指紋と版だけ出す。相手と突き合わせるときに使う")
    ap.add_argument("--expect",
                    help="**引き継ぎメモに書かれた指紋。**"
                         "手元のテンプレートと違えば異常終了する（終了コード 1）")
    ap.add_argument("--source", nargs="+",
                    help="正本（資料の HTML 一式）。メモに正本の指紋を書く／`--stale` で比べる")
    ap.add_argument("--stale",
                    help="**引き継ぎメモ（HANDOFF.md）。**書かれた正本の指紋と今の正本を比べ、"
                         "違えば「引き継ぎ一式は古い」と異常終了する。`--source` と一緒に使う")
    ap.add_argument("--snapshot", action="store_true",
                    help="**版を上げる前に、今の正本を退避する。**"
                         "snapshots/<日付>_<正本の指紋>/ に写す。`--source` と一緒に使う")
    args = ap.parse_args()

    if args.snapshot:
        # **正本が消えたときに、作り直しに頼らない**
        import shutil
        if not args.source:
            sys.exit("--snapshot には --source（今の正本）が要る")
        fp = source_fingerprint(args.source)
        dest = Path("snapshots") / f"{date.today().isoformat()}_{fp}"
        if dest.exists():
            print(f"既に退避済み: {dest}（正本は前回の退避から変わっていない）")
            return 0
        dest.mkdir(parents=True)
        for f in args.source:
            shutil.copy2(f, dest / Path(f).name)
        print(f"退避した: {dest}/（{len(args.source)} ファイル）")
        return checked.summary("handoff --snapshot", len(args.source), "ファイル", 0,
                               {"退避先": str(dest)})

    if args.stale:
        # **正本を直したら、引き継ぎ一式は古い。**機械で出す
        import re as _re
        if not args.source:
            sys.exit("--stale には --source（今の正本）が要る")
        memo = Path(args.stale).read_text(encoding="utf-8")
        m = _re.search(r"正本の指紋:\s*\**\s*([0-9a-f]{12})", memo)
        if not m:
            print("**メモに正本の指紋が無い。**古いかどうか判定できない。"
                  "`--source` を付けてメモを作り直す")
            return checked.summary("handoff --stale", 0, "件", 0,
                                   {"メモ": args.stale})
        want, got = m.group(1), source_fingerprint(args.source)
        ok = want == got
        print(f"メモの正本の指紋 {want} / 今の正本 {got}")
        print("**一致。**引き継ぎ一式は今の正本と同じ" if ok else
              "\n**引き継ぎ一式は古い。**正本がメモを書いた後に変わっている。"
              "\n渡す前に作り直す（HTML・md・メモを今の正本から書き出し直す）")
        code = checked.summary("handoff --stale", 1, "件", 0 if ok else 1,
                               {"メモ": want, "正本": got})
        return code or (0 if ok else 1)

    tpl = Path(args.template)
    if not tpl.is_dir():
        sys.exit(f"ディレクトリではない: {tpl}")

    got = fingerprint(tpl)

    if args.expect:
        # **規約に「突き合わせる」と書くだけでは守られない。**機械で止める。
        # 違う版で出しても、変換も検査も最後まで走りきってしまう
        want = args.expect.strip()
        ok = got == want
        print(f"{tpl.name}  指紋 {got}  引き継ぎメモ {want}")
        if ok:
            print("**一致した。**そのまま進めてよい")
        else:
            print("\n**指紋が違う。このまま変換しない。**")
            print("| | |")
            print("|---|---|")
            print(f"| 手元のテンプレート | {got} |")
            print(f"| 受け取った資料が前提にしている版 | {want} |")
            print("\n**名前が同じでも中身が違う。**どちらが新しいかを確かめる。")
            print("受け取った HTML のほうが新しいなら、"
                  "`gen_template.py` で作り直してから変換する。")
        code = checked.summary("handoff --expect", 1, "件", 0 if ok else 1,
                               {"手元": got, "メモ": want})
        # **不一致は異常終了させる。**サマリ行は「見た」ことしか示さない。
        # 終了コードで落とさないと、後続の変換がそのまま走る
        return code or (0 if ok else 1)

    if args.fingerprint:
        print(f"{tpl.name}  指紋 {got}  スキルの版 {skill_version()}")
        print("**両方が一致して初めて「同じ」。** 違えば新しい方に揃える")
        return 0

    size = "<manifest に無い>"
    man = tpl / "template-manifest.md"
    if man.exists():
        for line in man.read_text().splitlines():
            if line.startswith("slide_size:"):
                size = line.split(":", 1)[1].strip()
                break

    text = TEMPLATE.format(today=date.today().isoformat(), version=skill_version(),
                           name=tpl.name, size=size, fp=fingerprint(tpl),
                           src_fp=(source_fingerprint(args.source) if args.source
                                   else "<--source で正本を渡すと入る>"))
    if args.out:
        p = Path(args.out)
        if p.exists():
            sys.exit(f"既にある: {p}\n**上書きしない。** 別名にするか、手で追記する")
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text)
        print(f"書き出した: {p}")
    else:
        print(text)

    print("\n**<> の欄は人が書く。** 機械で決まるのは版・指紋・サイズだけ。")
    print("**「未決」と「触っていないもの」を空にしない。**"
          "無いと、受け取った側は全部を疑って読み直すことになる。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

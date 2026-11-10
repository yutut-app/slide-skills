#!/usr/bin/env python3
"""上書き・削除の前に、**利用者の手直しを差分で確かめて退避する。**

    python3 scripts/guard_edit.py record <自分が出した>.pptx       # 出すたびに控える
    python3 scripts/guard_edit.py check  <上書き・削除する>.pptx    # その前に必ず通す

`check` は次の順で動く。

1. 対象を `<案件>/work/backup/<日付>_<指紋>/` へ**退避する**（必ず。差分の有無に依らない）
2. 自分が最後に出した控え（`record` の写し）と比べる
3. **1枚でも変わっていれば止める**（終了コード1）。どの枚の何が変わったかを出す
4. 控えが無ければ**止める**。編集されたか分からないときは、手直しがある前提で扱う

**止まったら、上書き・削除しない。**手直しを残したまま、必要な枚だけ直す
（`skills/slider-pptx-craft/references/55_revise.md` の差分修正）。どうしても上書きするなら、利用者に差分を見せて確かめる。

実績として、別の版を削除するときに手直しの有無を確かめず、利用者が直した枚が戻った。
**確かめたかどうかは、言うだけでは証明にならない。**この道具の出力を報告に貼る。
"""

import argparse
import hashlib
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import checked
from diff_slides import canon, slide_names, texts
from paths import project_root


def fingerprint(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()[:8]


def changed_slides(ours: Path, target: Path):
    """[(枚, 種類, 中身)]。比べ方は diff_slides.py と同じ（正規化した XML と文字）。"""
    a, b = zipfile.ZipFile(ours), zipfile.ZipFile(target)
    na, nb = slide_names(a), slide_names(b)
    out = []
    for i in range(max(len(na), len(nb))):
        no = i + 1
        if i >= len(na):
            out.append((no, "利用者が追加した枚", ""))
            continue
        if i >= len(nb):
            out.append((no, "利用者が消した枚", ""))
            continue
        ba, bb = a.read(na[i]), b.read(nb[i])
        if ba == bb or canon(ba) == canon(bb):
            continue
        ta, tb = texts(ba), texts(bb)
        if ta != tb:
            removed = [x for x in ta if x not in tb][:2]
            added = [x for x in tb if x not in ta][:2]
            what = "、".join(filter(None, [
                ("消: " + " / ".join(x[:16] for x in removed)) if removed else "",
                ("足: " + " / ".join(x[:16] for x in added)) if added else ""]))
            out.append((no, "文字が変わった", what))
        else:
            out.append((no, "見た目が変わった", "位置・色・書式のいずれか"))
    return out


def main():
    ap = argparse.ArgumentParser(description="上書き・削除の前に、手直しを確かめて退避する")
    ap.add_argument("mode", choices=["record", "check"])
    ap.add_argument("pptx")
    ap.add_argument("--ours", help="比べる控え。省くと work/backup/ours/<同じ名前>")
    ap.add_argument("--project", help="案件のディレクトリ。省くと project.md を上に辿る")
    args = ap.parse_args()

    target = Path(args.pptx).expanduser().resolve()     # **最初に絶対パスへ**
    if not target.exists():
        print(f"**対象が無い：** {target}")
        return checked.summary(f"guard_edit {args.mode}", 0, "ファイル", 1, {})
    root = project_root(args.project, hint=str(target))
    ours_dir = root / "work" / "backup" / "ours"

    if args.mode == "record":
        ours_dir.mkdir(parents=True, exist_ok=True)
        dest = ours_dir / target.name
        shutil.copy2(target, dest)
        print(f"控えた：{dest.relative_to(root)}（指紋 {fingerprint(target)}）")
        return checked.summary("guard_edit record", 1, "ファイル", 0, {"案件": root.name})

    # 1. 退避（必ず）
    bk = root / "work" / "backup" / f"{date.today().isoformat()}_{fingerprint(target)}"
    bk.mkdir(parents=True, exist_ok=True)
    if not (bk / target.name).exists():
        shutil.copy2(target, bk / target.name)
    print(f"退避した：{(bk / target.name).relative_to(root)}")

    # 2. 控えと比べる
    ours = Path(args.ours).expanduser().resolve() if args.ours else ours_dir / target.name
    if not ours.exists():
        print("\n**比べる控えが無い。止める。**編集されたか分からないので、手直しがある前提で扱う。")
        print("上書き・削除せず、必要な枚だけ直す（`skills/slider-pptx-craft/references/55_revise.md` の差分修正）。")
        print("次からは、自分が出すたびに `guard_edit.py record <出した pptx>` で控える。")
        code = checked.summary("guard_edit check", 1, "ファイル", 1, {
            "対象": target.name, "控え": "**無い**", "退避": "済み"})
        return code or 1

    diffs = changed_slides(ours, target)
    if diffs:
        print(f"\n**利用者の手直しがある（{len(diffs)} 枚）。上書き・削除しない。**\n")
        print("| 枚 | 種類 | 中身 |")
        print("|---|---|---|")
        for no, kind, what in diffs:
            print(f"| {no} | {kind} | {what} |")
        print("\n手直しを残したまま、必要な枚だけ直す。上書きするなら、**この表を利用者に見せて確かめる。**")
    else:
        print("\n手直しは無い（控えと同じ）。退避済みなので、上書きしてよい。")

    code = checked.summary("guard_edit check", 1, "ファイル", len(diffs), {
        "対象": target.name, "控え": ours.name, "退避": "済み",
        "変わった枚": f"{len(diffs)} 枚"})
    return code or (1 if diffs else 0)


if __name__ == "__main__":
    sys.exit(main())

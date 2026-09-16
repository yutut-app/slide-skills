# slide-skills

日本語の報告・提案スライドを作るためのスキル。

## 収録スキル

| スキル | 役割 | 成果物 |
|---|---|---|
| `skills/slider-craft` | 資料の**中身**を作る。意図の確定、骨子、図解の型、日本語の文体、検査 | **HTML 一式** |
| `skills/slider-deck` | HTML を **pptx にする**。変換、溢れの検査、ずれの差し戻し | pptx ＋ 指摘 |
| `skills/slider-reskin` | HTML の**器**を付け替える。別テンプレートへの載せ替え | 別テンプレの HTML |

### 機械ごとに、要るスキルは違う

| 機械 | 入れるスキル | やること |
|---|---|---|
| **作る側** | craft ＋ deck ＋ reskin | 中身を作り、テンプレートを保守し、配る |
| **提出先・第三者** | **deck だけ** | 受け取った HTML を pptx にし、見た目を検査して返す |

**`slider-deck` は単独で完結する。** 手順の中で他のスキルのファイルを読まない。
規約の中で「配布元の規約」と書いてあるものは**手元に無くてよい。**
そこはこのスキルの担当ではないので、**気づいたことは報告に書いて返す。**

**HTML が正本、pptx は派生物。** pptx を手で直しても次の変換で消える。
直すのは常に HTML。

## 必要なもの

### 共通

```
python3 -m pip install python-pptx lxml openpyxl
```

`lxml` は `python-pptx` が連れてくるが、**明示して入れる。**
依存の都合で外れると、変換が丸ごと動かなくなる。

**画像化の道具がどちらか1つ要る。**無いと見た目の検査ができない。

| 道具 | 何に使うか |
|---|---|
| **Edge か Chrome** | HTML を絵にする（`html_png.py`）。**Windows は Edge が標準で入っている** |
| **LibreOffice か PowerPoint** | pptx を絵にする（`qa_render.py`） |

### Windows

```
py -m pip install python-pptx lxml openpyxl pywin32
```

**`pywin32` を入れると、PowerPoint COM で pptx を絵にできる。**
LibreOffice を入れる必要はない。**Edge は標準で入っているので追加は要らない。**

**Chrome は要らない。**既定の場所に無いときは `--browser` で実行ファイルを渡す。

**書体は Verdana とメイリオ。**どちらも標準で入っている。
**書体が無いと折り返し位置が変わり、検査が当てにならない。**

### macOS

```
python3 -m pip install python-pptx lxml openpyxl
brew install --cask libreoffice
```

LibreOffice は PATH に入らない。スクリプト側で絶対パスを見ている。

**メイリオは標準では入っていない。**検査の1行の `書体=` 欄に不足と出たら、
**折り返し位置は提出先の環境と違う。**そのまま「確認した」と書かない。

### 入る前に確かめる

```
python3 skills/slider-deck/scripts/qa_render.py --probe
python3 skills/slider-craft/scripts/html_png.py --probe
```

**無い道具があるまま進めない。**「検査した」と書けなくなる。

## 登録

実体はこのリポジトリに置き、`~/.claude/skills/` から参照する。

```bash
# macOS / Linux
for s in slider-craft slider-deck slider-reskin; do
  ln -sfn "$PWD/skills/$s" ~/.claude/skills/$s
done
```

```powershell
# Windows (管理者の PowerShell)
foreach ($s in "slider-craft","slider-deck","slider-reskin") {
  New-Item -ItemType SymbolicLink -Force `
    -Path "$env:USERPROFILE\.claude\skills\$s" -Target "$PWD\skills\$s"
}
```

確認は一覧を見る。**リンク切れは一覧から無言で消えるだけ**なので、
スキルが出てこないときは最初にここを見る。

## 資料テンプレート

**`assets/templates/deck/<テンプレ名>/` に置く。** 3つのスキルが同じ場所を見る。
**2箇所に置くと必ず食い違う。**

```
assets/templates/deck/<テンプレ名>/
├ 001.html 〜 004.html / free.html   レイアウトごとの HTML
├ print.html                          全部を並べた閲覧用
├ template-manifest.md                **px→pt の係数。無いと寸法が狂う**
├ images/                             ロゴなど
└ source.pptx                         実物（任意。読むだけ）
```

**テンプレートは git に入れていない**（利用者のデータ）。
**クローン直後には `assets/templates/deck/` が無い。**
使う機械には、**あらかじめ置いておく。**置いていない機械で急ぎ使うときだけ、
受け取った資料の HTML から作り直せる（`slider-deck/scripts/gen_template.py`）。
**ロゴは作り直せない**ので、画像は HTML と一緒に受け取る必要がある。

別の場所に置くなら環境変数を向ける。

```
Windows (PowerShell):  $env:SLIDE_TEMPLATE_DIR = "D:\templates"
macOS:                 export SLIDE_TEMPLATE_DIR=~/templates
```

置いたら、**一覧に出ることを確かめてから始める。**

```
python3 skills/slider-reskin/scripts/reskin_html.py --list
```

**名前が同じでも中身が違うことがある。**指紋を突き合わせる。

```
python3 skills/slider-craft/scripts/handoff.py <テンプレ> --fingerprint
```

テンプレートが1つも無いときは、汎用のものを作れる（**同梱していない**）。

```
python3 skills/slider-craft/scripts/gen_deck_templates.py
```

## 更新のしかた

**このリポジトリは配布物。** 更新するたびにコミットが積まれるので、
**どのファイルが変わったかは履歴で追える。**

```bash
git log --oneline          # 何を変えたか
git log --stat -1          # 直近でどのファイルが変わったか
git pull                   # 受け取った側の更新
```

**出してはいけないものが混ざっていないかは、公開の前に機械で見ている。**
ファイルの中身と**コミット文の両方**を検査し、1件でも当たれば push しない。

---

## 置き場所の決まり ★

**スキルの資産はスキルの中。利用者のデータはスキルの外。**

スキルは `~/.claude/skills/<name>` から参照される。
**その外に置いたものは、スキルの一部ではない。**

| | 置き場所 | 理由 |
|---|---|---|
| **図の雛形25件** | `skills/slider-deck/assets/templates/` | 配布物。**スキルだけ配っても付いてくる** |
| 図の索引（選び方） | 各スキルの `assets/templates/INDEX.md` | craft は型を選び、deck は作る。**直すときは両方直す** |
| **資料テンプレート** | `assets/templates/deck/`（リポジトリ直下） | **利用者のデータ。**git に入れない。3スキルが同じものを見る |

リポジトリ直下に置いたところ、**スキルだけを配ると雛形が0件**になり、
規約が書いている `assets/templates/INDEX.md` も空振りした。

### 図の雛形

25件（グラフ11・QC7つ道具7・新QC7つ道具7）。**クローンすれば手元にある。**

```
skills/slider-deck/assets/templates/
├ INDEX.md    索引と選び方
├ charts/     グラフ（Excel）
├ qc7/        QC7つ道具
└ n7/         新QC7つ道具
```

## 主な道具

| 道具 | いつ | スキル |
|---|---|---|
| `make_html.py` | .pptx テンプレートから編集用 HTML を起こす | craft |
| `html_png.py` | **HTML を絵にする。**利用者が見ている絵そのもの | craft |
| `extract_assets.py` | 実物の .pptx から画像（ロゴ）を取り出す | craft |
| `handoff.py` | 引き継ぎメモ。`--fingerprint` で版と指紋 | craft |
| `html2pptx.py` | **HTML → pptx** | deck |
| `stage0.py` | **段階0。**1枚だけ変換してテンプレート自体のずれを潰す | deck |
| `fit_check.py budget` | **文言を書く前。**入る文字数と縦グリッド | deck |
| `fit_check.py check` | 作った後。重なり・はみ出し・溢れ | deck |
| `qa_render.py` | pptx / xlsx を1枚1枚 PNG に | deck |
| `diff_slides.py` | 手直しの後。変えたスライドを機械的に出す | deck |
| `pptx_diff_html.py` | 手で直した pptx と HTML の差 | deck |
| `reskin_html.py` | 別テンプレートへの載せ替え | reskin |

**検査の道具は必ず1行を出す。**

```
検査: <道具>  対象 n<単位>  指摘 n件  前提: ...
```

**対象が0なら異常終了する**（終了コード 2）。
「指摘ゼロ」と「何も見ていない」を出力で区別するため。
**報告にはこの行をそのまま貼る。**「確認しました」は証拠にならない。

## スクリプトの重複

`read_template_style.py` は3つのスキルに同じものを置いている。
`checked.py` `html_png.py` `handoff.py` も同様。スキルを単独で置いても動く必要があるため。
**直すときは全部直す。**

## 資料を外に出さない

扱うのは、まだ外に出ていない資料。**一度外に出ると取り消せない。**

- **公開・リンク共有・第三者が開ける状態にしない**
- **資料とテンプレートを git に入れない**（`.gitignore` で除外）
- オンライン変換サービスに上げない

ローカルにファイルを書くのは問題ない。**「利用者に見せる」と「外部に載せる」を混同しない。**

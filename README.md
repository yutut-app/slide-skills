# slide-skills

日本語の報告・提案スライドを作るためのスキル。

## 案件で使うとき

**`docs/handoff-to-project.md` を最初に読む**（このリポジトリの手元の写しにだけある。
配布物には含まれない）。残件・必ず通す検査・踏んだ落とし穴が1本にまとまっている。

**案件のフォルダの `README.md` にも、そのパスを1行書く。**
案件側のセッションは案件フォルダとスキルしか見ないので、書かないと辿れない。

## 収録スキル

**作るもの（HTML／pptx）と工程で5つに分ける。**呼び分けの正本は `assets/content/00_route.md`。
どれが先に呼ばれても、最初の判定表（`scripts/entry_check.py`）で担当が決まる。

| スキル | 役割 | 成果物 |
|---|---|---|
| `skills/slider-html-craft` | 資料の**中身**を作り HTML にする。新規、または前回の会議の資料から | **HTML 一式** |
| `skills/slider-html-reskin` | html-craft の HTML を**器**（テンプレート・客先 pptx）に載せて **pptx にする** | pptx |
| `skills/slider-pptx-craft` | HTML を通さず **pptx を直接作り、直す**。新規、または前回の会議の pptx から。人の手直しを消さない | pptx |
| `skills/slider-deck` | できた pptx を**渡す前に検査する**。直さない | 指摘と差し戻し先 |
| `skills/slider-script` | **カンペ**と pptx の**ノート**を作る | `script.md`、ノート入りの pptx |

```
HTML の道:  slider-html-craft → slider-html-reskin → slider-deck
pptx の道:  slider-pptx-craft → slider-deck
```

中身の規約（意図・骨子・図・日本語）は2つの craft が共通で読むので、スキルのフォルダではなく
**リポジトリ直下の `assets/content/`** に1つだけ置く。**リポジトリごと clone して使う。**

### 機械ごとに、要るスキルは違う

| 機械 | 入れるスキル |
|---|---|
| **作る側** | 5つ全部 |
| **提出先・第三者** | 受け取った HTML を pptx にするなら html-reskin ＋ deck。検査だけなら deck |

## 登録と更新

実体はこのリポジトリに置き、`~/.claude/skills/` から参照する。**clone した機械では、`git pull` の後に毎回これを1回実行する。**
スキルの名前を変えたとき、古いリンクを外して新しいリンクを張る（中身は消さない）。

```bash
git pull
sh tools/link.sh            # 登録する。旧名（slider-craft・slider-reskin・create-slide）のリンクを外す
sh tools/link.sh --check    # 何もせず、状態だけ出す
```

**版の確かめ方：**どのスキルでも、最初の判定表の1行目に「いま動いているスキル（版、日付）」が出る。
レビューや不具合の報告には、その行を貼る。**古い版が動いていたら、`git pull` と `sh tools/link.sh` から。**

Windows はリンクの代わりにフォルダを写す（管理者の PowerShell で `New-Item -ItemType SymbolicLink` でもよい）。
写した場合は `assets/content/` が届かないので、リポジトリごと置き、`skills\<名前>` を指すリンクを張る。

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
受け取った資料の HTML から作り直せる（`slider-html-reskin/scripts/gen_template.py`）。
**ロゴは作り直せない**ので、画像は HTML と一緒に受け取る必要がある。

別の場所に置くなら環境変数を向ける。

```
Windows (PowerShell):  $env:SLIDE_TEMPLATE_DIR = "D:\templates"
macOS:                 export SLIDE_TEMPLATE_DIR=~/templates
```

置いたら、**一覧に出ることを確かめてから始める。**

```
python3 skills/slider-html-reskin/scripts/reskin_html.py --list
```

**名前が同じでも中身が違うことがある。**指紋を突き合わせる。

```
python3 skills/slider-html-craft/scripts/handoff.py <テンプレ> --fingerprint
```

テンプレートが1つも無いときは、汎用のものを作れる（**同梱していない**）。

```
python3 skills/slider-html-craft/scripts/gen_deck_templates.py
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
| **図の雛形25件** | `skills/slider-html-reskin/assets/templates/` | 配布物。**スキルだけ配っても付いてくる** |
| 図の索引（選び方） | 各スキルの `assets/templates/INDEX.md` | craft は型を選び、deck は作る。**直すときは両方直す** |
| **資料テンプレート** | `assets/templates/deck/`（リポジトリ直下） | **利用者のデータ。**git に入れない。3スキルが同じものを見る |

リポジトリ直下に置いたところ、**スキルだけを配ると雛形が0件**になり、
規約が書いている `assets/templates/INDEX.md` も空振りした。

### 図の雛形

25件（グラフ11・QC7つ道具7・新QC7つ道具7）。**クローンすれば手元にある。**

```
skills/slider-html-reskin/assets/templates/
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
| `handoff.py` | 引き継ぎメモ。`--fingerprint` で版と指紋、**`--expect` で違えば止める** | craft / deck |
| `html2pptx.py` | **HTML → pptx**（位置は既定で組み直す。写すなら `--keep-coords`） | deck |
| `stage0.py` | **段階0。**1枚だけ変換してテンプレート自体のずれを潰す | deck |
| `fit_check.py budget` | **文言を書く前。**入る文字数と縦グリッド | deck |
| `fit_check.py check` | 作った後。重なり・はみ出し・溢れ | deck |
| `qa_render.py` | pptx / xlsx を1枚1枚 PNG に | deck |
| `diff_slides.py` | 手直しの後。変えたスライドを機械的に出す | deck |
| `pptx_diff_html.py` | 手で直した pptx と HTML の差 | deck |
| `chart_xlsx.py` | **HTML の図から Excel を作り、pptx に貼る** | deck |
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

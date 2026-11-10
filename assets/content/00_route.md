# 00_route — 5つのスライドスキルの呼び分け

**5スキルのどれから始まっても、最初に `scripts/entry_check.py` を通し、出た表で担当を決める。**
ここはその判定の考え方の正本。**各スキルの本文に写さない。**

## 分け方の軸

**何を作るか（HTML／pptx）と、工程（作る／変換する／検査する／話す）。**
実績: 以前は「どの層に手を入れるか（中身／器／ファイル）」で分けていたが、pptx が正本の案件で
「ファイルを直す」と「中身を作る」が1つの作業になり、どのスキルにも当てはまらなかった。

| スキル | 作るもの | 工程 | 入力 | 出力 |
|---|---|---|---|---|
| `slider-html-craft` | HTML | 作る（新規・前回の会議の資料から） | 依頼・メモ・前回の資料 | HTML 一式 |
| `slider-html-reskin` | pptx | 変換する（器に載せる） | html-craft の HTML | pptx |
| `slider-pptx-craft` | pptx | 作る・直す（新規・前回の会議の pptx から） | テンプレ・前回の pptx | pptx |
| `slider-deck` | — | 検査する（直さない） | html-reskin / pptx-craft の pptx | 指摘と差し戻し先 |
| `slider-script` | script.md・ノート | 話す内容を作る | スライド | 台本・ノート入りの pptx |

```
HTML の道:  html-craft → html-reskin → deck
pptx の道:  pptx-craft → deck
どちらも:   → script（カンペ・ノート）→ deck（ノートの検査）
```

## 依頼の言葉から担当を決める

| 依頼 | 担当 | `--want` |
|---|---|---|
| 「HTMLで作って」「構成を考えて」「このHTMLを直して」 | html-craft | html |
| 「HTMLをパワポにして」「pptxに変換」「テンプレート／客先書式に載せて」（HTML がある） | html-reskin | convert |
| 「パワポで作って」「前回の会議資料をベースに」「このパワポを直して」「Rev3 で上書き」 | pptx-craft | pptx |
| 「検査して」「崩れていないか」「渡す前に確認」「差分を出して」 | deck | check |
| 「カンペ」「台本」「ノートに入れて」 | script | script |

**言葉で決まらないとき**（「資料作って」「パワポ作って」「自作パワポスキルで」）は、手元の材料で決める。

| 手元 | 担当 |
|---|---|
| pptx だけ（前回の会議の資料など） | pptx-craft（検査だけなら deck） |
| HTML だけ | **聞く**：中身を直す（html-craft）か、pptx にする（html-reskin）か |
| 両方 | **聞く**：どちらが正本か |
| 何も無い | **聞く**：HTML で作るか、pptx で作るか（外すと全工程がやり直し） |

**聞くのは判定表で1回だけ。**版の扱いなど、他の未決とまとめて聞く。

## 隣のスキルが無い環境

判定表は、5スキルの登録の有無と版を出す。**担当のスキルが未登録なら、その旨を報告して止まる。**
代わりに自分で担当外の作業をしない（実績: 担当でないスキルが呼ばれ、使っていない手順を「使った」と記録した）。
登録は `sh tools/link.sh`（clone したリポジトリで `git pull` の後に毎回）。

## 旧スキル

| 旧名 | 後継 |
|---|---|
| slider-craft | slider-html-craft |
| slider-reskin | slider-html-reskin（HTML → pptx）、slider-pptx-craft（pptx → pptx） |
| create-slide | slider-html-craft、slider-pptx-craft |

**判定表に「旧スキルの登録が残っている」と出たら、使わない。**外すのは `sh tools/link.sh`（中身は消えない）。

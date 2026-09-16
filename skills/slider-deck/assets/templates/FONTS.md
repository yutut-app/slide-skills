# 図テンプレートの書体

`charts/` `qc7/` `n7/` のテンプレートは、**資料テンプレートに合わせて再生成する。**
図だけ別の書体だと、資料に貼った瞬間に浮く。

## 現在の設定

| 用途 | 書体 |
|---|---|
| 日本語（ea） | Meiryo UI |
| 欧文（latin / cs） | Verdana |

`assets/templates/deck/` に置いた資料テンプレートの実体に合わせた値。
**テーマの指定（`+mj-*` / `+mn-*`）ではなく、スライドで実際に使われている書体**を採った。
テーマに日本語（ea）が未指定で、Aptos などの欧文だけが入っていることがあるため。

## 再生成

```bash
cd skills/slider-craft
python3 scripts/gen_chart_templates.py --font "Meiryo UI" --latin "Verdana"
python3 scripts/gen_qc7_templates.py   --font "Meiryo UI" --latin "Verdana"
python3 scripts/gen_pptx_templates.py  --font "Meiryo UI" --latin "Verdana"
```

資料テンプレートが変わったら、まず書体を調べ直す。

```bash
python3 skills/slider-reskin/scripts/read_template_style.py <テンプレ>.pptx
```

テーマの日本語が未指定なら、**実際に使われている書体**（出力の「実際に使われている
フォント」）から決める。

## 注意

**書体名を指定しても、その書体が入っていない環境では代替表示になる。**
Meiryo UI は Windows には標準で入るが、macOS には無い。
指定した書体が無い環境では代替書体で描かれるため、**文字のはみ出しは厳しめに見る**
（`slider-craft/references/60_qa.md`）。

# visual_check の回帰見本

**検査を直したら、両方を通す。**

```bash
python3 ../visual_check.py violations.html --manifest template-manifest.md   # 指摘4件・終了1
python3 ../visual_check.py clean.html      --manifest template-manifest.md   # 指摘0件・終了0
```

`violations.html` には、実運用で出た5つの症状を1枚に入れてある。
下端のずれ／小さい文字／文字の重なり／枠に入らない値／同じ数値の近接した二重。

**見本だけで通る検査にしない。**実案件の HTML でも1件通す。
実案件で誤検出が出たら、**資料を直すのではなく検査を直す**（実際に、透かしの帯と
軸ラベルの接触で40件以上の誤検出が出て、共通パーツの扱いと重なりの下限を入れた）。

`chart_no_values.html` は、グラフに数（`data-values`）が無い枚と、ある枚を1枚に入れてある。
**無い方だけが要確認に出る**（指摘0件・終了0。数が無いことは、人が判断する）。

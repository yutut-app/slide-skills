# layout_check の回帰見本

**見本をリポジトリに置かない**（案件の形そのものなので、作業スペースに作って通す）。

```bash
WS="<作業スペース>/YYmmdd_layout-test"
mkdir -p "$WS/in" && printf '# test\n' > "$WS/project.md"
python3 ../layout_check.py "$WS"                      # 指摘0件・終了0

mkdir -p "$WS/qa_html" "$WS/out/deck_20260914-1649"   # 散らばらせる
python3 ../layout_check.py "$WS"                      # 指摘3件・終了1
```

**実案件で出た症状を3つ入れてある。**直下の別フォルダ、同じ用途の別名、
版の名前が決めごとと違う（日時入りは要確認で出る）。

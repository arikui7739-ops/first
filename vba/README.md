# VBAモジュール（UTF-8ソース）

ここにあるファイルはUTF-8で保存されたソース（Git管理・差分確認用）。

**VBEにインポートする前に、必ずShift-JIS(CP932)に変換すること。** UTF-8のままインポートすると文字化けする。

```python
with open('Module3_new.bas', encoding='utf-8') as f:
    content = f.read()
with open('Module3_new_sjis.bas', 'w', encoding='cp932', errors='replace') as f:
    f.write(content)
# 往復検証
with open('Module3_new_sjis.bas', encoding='cp932') as f:
    conv = f.read()
assert content == conv, "round-trip mismatch — mojibake risk"
```

各モジュールの内容・KPI計算式・修正履歴はリポジトリルートの `CLAUDE.md` を参照。

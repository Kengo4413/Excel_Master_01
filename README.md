# Excel_Master_01

Excel 実務スキルを体系的に学べる演習教材です。関数の解説マニュアルと、実際に手を動かせる演習データ（`.xlsx`）をセットで提供します。

## 内容

| ファイル | 説明 |
|----------|------|
| `excel.md` | Excel 実務完全マニュアル（見積・統合・管理・効率化） |
| `excel_演習_問題と解答.md` | 各関数に対応した演習問題・解答・解説 |
| `excel_演習_データ.xlsx` | 演習用 Excel ファイル（6シート構成） |
| `build_excel_practice.py` | xlsx を再生成する Python スクリプト |
| `excel.html` / `excel_演習_問題と解答.html` | HTML 版（ブラウザ閲覧用） |

## 学べる関数・機能

| カテゴリ | 関数・機能 |
|----------|-----------|
| 集計 | SUM, SUMIF, SUMIFS, SUMPRODUCT, SUBTOTAL |
| 端数処理 | ROUND, ROUNDDOWN, ROUNDUP |
| 検索・参照 | VLOOKUP, XLOOKUP, IFERROR |
| 文字列操作 | LEFT, RIGHT, MID, LEN, TRIM, SUBSTITUTE |
| カウント | COUNTIF, COUNTIFS |
| 日付・営業日 | WORKDAY, NETWORKDAYS, EOMONTH, DATEDIF |
| 実務 Tips | 入力規則（プルダウン）, 絶対参照（F4）, テーブル化（Ctrl+T） |

## 使い方

1. `excel_演習_データ.xlsx` を Excel で開く
2. `excel_演習_問題と解答.md`（または HTML 版）を見ながら各シートの問題に取り組む
3. 解答・解説で理解を確認

## xlsx の再生成

```bash
PYTHONPATH="./.pydeps" python3 build_excel_practice.py
```

`openpyxl` が未インストールの場合:

```bash
pip install --target ./.pydeps openpyxl
```

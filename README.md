# MeGov
eGov で配布されている法令 XML ファイルを Word（.docx）に変換する Python コード
## 💬ストーリー
法改正等の度に法令をコピーして体裁を整えたり…、日本中からこの面倒な作業を無くそうと思い作成しました。
## 🔍Word（.docx）ファイルの探し方
1. このページの「Go to file」を選択する
2. 「WeGov/」に続いて法令のキーワードを入力して検索する
3. 目的の法令ファイルを選択する
4. 「Download」を選択してダウンロードする
## 🙇免責事項
- 図表の抽出など一部機能が実装されていません。コード更新次第、Word 文書も随時更新致します。
- 私は本業がプログラマーではなく、素人コードです。
# Python コード
## 🐍使い方
python3 wegov.py eGovXML.xml exportDirectory
## 📚ライブラリ
- 🥣BeautifulSoup4(bs4)
- 📝python-docx
# 関連レポジトリ
- [MeGov](https://github.com/fts141/MeGov)
eGov で配布されている法令 XML ファイルを Markdown（.md）に変換する Python コード
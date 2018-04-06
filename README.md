# リライトシート　検索ランキング　自動ツール

## setup
- `pip install -r requirements.txt`
- https://developers.google.com/sheets/api/quickstart/python?authuser=3 の Step 1 を行う
- 最終的にはOAuth 2.0 クライアント ID を作成し、`client_secrets.json`というファイルに保存する

## Running the script
- `python get_search_rank.py`
- https://docs.google.com/spreadsheets/d/1wE2YkNIT7wYuR7Ul23Xv98WqIQwe0gNnOuw3a-3jJpY/edit#gid=588079040
- スクリプトの大まかな流れとしては
1. Google spreadsheet API を使って上記シートの「リライトシート」タブからキーワード一覧を読見込む
2. Google search console API を使って各キーワードの過去一週間の平均ランキングを抽出
3. spreadsheet APIで新しいcolumnに書き込む

大体１〜２分ぐらいかかります。

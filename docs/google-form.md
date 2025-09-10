# Google form 関連のメモ

Google Form は構成は、ざっと以下のようになっています。

- フォーム全体の設定
  - タイトル
  - 説明
- 実施者向けの説明
  - type: SECTION_HEADER
  - helpText: 実施者向けの説明
- 見積もり対象 1
  - 説明
    - type: PAGE_BREAK
    - helpText: 回答者への作業指示
  - （設問）見積もりの前提、質問
    - type: PARAGRAPH_TEXT
  - （設問）見積もり値
    - type: LIST
    - choices: 1,2,3,5,8,13,21,34,55,89,skip
- 見積もり対象 2
- ...

ログからも判断できると思います。

```plaintext
12:13:16    お知らせ    実行開始
12:13:16    情報    [INFO] runDebugForm start
12:13:16    情報    [INFO] Loaded table 見積もり必要_デバッグ {"a1":"見積もり設定!B51:C53","countRows":1,"keys":1,"tableId":"738923658"}
12:13:16    情報    [INFO] Logging form {"formUrl":"https://docs.google.com/forms/d/14HRtRyqyXXzbjBakojxwaeHdWNNgu-s1ym4GcVn5DbE/edit"}
12:13:17    情報    [INFO] Form info {"title":"yyyy-mm-dd async ポーカー","description":"同期的に時間をとりづらいチームのために、非同期でポーカーを実施させていただきます。\n\n本日の 16:00 が締め切りです。ご対応よろしくお願いいたします。"}
12:13:17    情報    [INFO] Form item {"index":0,"type":"SECTION_HEADER","title":"ポーカー Google Form の作成者向け説明","helpText":"※ 回答者の方は、こちらをお読みいただく必要はありません。次にお進みください。\nポーカーは GAS で作成します。\n当日に出勤している方のみを対象としてください。\n作成したら Slack で実施を依頼してください。\n回答が集まったら、結果を結果スプシにコピペして、 Slack で確認依頼してください。\n確定した結果を issue の estimate に入れていきます。"}
12:13:17    情報    [INFO] Form item {"index":1,"type":"PAGE_BREAK","title":"https://github.com/xxx/yyy/issues/111","helpText":"👆👆👆タイトルが issue へのリンクになっています。\n\n見積もりと、どういう作業を含めることを前提としたか入力してください。skip の場合は質問や理由を入力してください。"}
12:13:17    情報    [INFO] Form item {"index":2,"type":"PARAGRAPH_TEXT","title":"E1. 見積もりの前提、質問"}
12:13:18    情報    [INFO] Form item {"index":3,"type":"LIST","title":"E1. 見積り値","choices":["1","2","3","5","8","13","21","34","55","89","skip"]}
12:13:18    情報    [INFO] runDebugForm done {"ms":2136}
12:13:18    お知らせ    実行完了
```

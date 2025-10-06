# My Google Apps Script

- [My Google Apps Script](#my-google-apps-script)
  - [概要](#概要)
  - [フォルダ構成](#フォルダ構成)
  - [権限(Scopes)](#権限scopes)
  - [事前準備(共通設定)](#事前準備共通設定)
  - [共通ユーティリティ `google-spreadsheets/utils.gs`](#共通ユーティリティ-google-spreadsheetsutilsgs)
  - [スクリプト別 仕様](#スクリプト別-仕様)
    - [`api/tiktok.gs`](#apitiktokgs)
    - [`google-documents/import-calendar.gs`](#google-documentsimport-calendargs)
    - [`google-drive/create-folder.gs`](#google-drivecreate-foldergs)
    - [`google-forms/answer-limit.gs`](#google-formsanswer-limitgs)
    - [`google-forms/default-value-1-3.gs`](#google-formsdefault-value-1-3gs)
    - [`google-forms/google-forms-to-slack/ver-1.gs`](#google-formsgoogle-forms-to-slackver-1gs)
    - [`google-forms/google-forms-to-slack/ver-2.gs`](#google-formsgoogle-forms-to-slackver-2gs)
    - [`google-spreadsheets/convert-schedules.gs`](#google-spreadsheetsconvert-schedulesgs)
    - [`google-spreadsheets/plot-calendar.gs`](#google-spreadsheetsplot-calendargs)
    - [`google-spreadsheets/remind.gs`](#google-spreadsheetsremindgs)
    - [`google-spreadsheets/get_list.gs`](#google-spreadsheetsget_listgs)
    - [`google-spreadsheets/get-holiday.gs`](#google-spreadsheetsget-holidaygs)
    - [`google-spreadsheets/insert-updated-at.gs`](#google-spreadsheetsinsert-updated-atgs)
    - [`google-spreadsheets/detect-header-change.gs`](#google-spreadsheetsdetect-header-changegs)
    - [`google-spreadsheets/create-slack-channel.gs`](#google-spreadsheetscreate-slack-channelgs)
    - [`google-spreadsheets/prevent-archive.gs`](#google-spreadsheetsprevent-archivegs)
    - [`google-spreadsheets/morphological-analysis.gs`](#google-spreadsheetsmorphological-analysisgs)
    - [`google-spreadsheets/extract-keyword.gs`](#google-spreadsheetsextract-keywordgs)
    - [`google-spreadsheets/check-for-updates-with-redash-api.gs`](#google-spreadsheetscheck-for-updates-with-redash-apigs)
    - [`google-spreadsheets/detect-updating-sheets.gs`](#google-spreadsheetsdetect-updating-sheetsgs)
    - [`google-spreadsheets/convet-data-for-ai.gs`](#google-spreadsheetsconvet-data-for-aigs)
  - [推奨トリガー設定一覧](#推奨トリガー設定一覧)
  - [セキュリティ/運用上の注意](#セキュリティ運用上の注意)
  - [動作確認チェックリスト](#動作確認チェックリスト)
  - [ライセンス](#ライセンス)

## 概要

- 個人/業務で作成した Google Apps Script を用途別に整理したリポジトリです。
- スプレッドシート/フォーム/ドキュメント/ドライブ/外部 API(Slack・TikTok・Yahoo)連携のサンプルや運用スクリプトを含みます。

## フォルダ構成

- `api/`
  - `tiktok.gs`: TikTok oEmbed を用いてシートへ情報とサムネを挿入
- `google-documents/`
  - `import-calendar.gs`: Google ドキュメントに日付テンプレ/カレンダー予定を挿入
- `google-drive/`
  - `create-folder.gs`: クライアント/ブランド/案件用のフォルダ群作成・検索
- `google-forms/`
  - `answer-limit.gs`: 回答上限でフォーム受付を停止
  - `default-value-1-3.gs`: フォームのプレフィル URL 生成(複製用リンク)
  - `google-forms-to-slack/ver-1.gs`: フォーム回答を Slack に通知(単一チャンネル)
  - `google-forms-to-slack/ver-2.gs`: フォーム回答を Slack に通知(回答内容で送信先や文面を分岐)
- `google-spreadsheets/`
  - `utils.gs`: 共通ユーティリティ/設定(重要)
  - `convert-schedules.gs`: スケジュール表から「一覧」へ期日/フラグを自動反映
  - `plot-calendar.gs`: `INPUT` → `OUTPUT` へカレンダー可視化(祝日考慮)
  - `remind.gs`: PM/OPE/KOL 向け日次リマインドの Slack 通知
  - `get_list.gs`: 来月の薬事チェック案件リストを Slack に通知
  - `get-holiday.gs`: 祝日一覧を取得(カレンダー API)しシートへ出力
  - `insert-updated-at.gs`: 特定シートの更新行へ更新日を自動記入
  - `detect-header-change.gs`: ヘッダー名変更を検知して Slack 通知
  - `create-slack-channel.gs`: Slack チャンネル作成/招待/ブックマーク/Canvas 作成
  - `prevent-archive.gs`: チャンネルのアーカイブ防止メッセージを定期送信
  - `morphological-analysis.gs`: Yahoo 形態素解析 API で語彙集計
  - `extract-keyword.gs`: 正規表現による PR/AF/ORG/ALL のハッシュタグ/語彙集計
  - `check-for-updates-with-redash-api.gs`: 指定シートの更新検知(差分抽出/LCS)と Slack 通知
  - `detect-updating-sheets.gs`: シンプル版の更新検知と Slack 通知
  - `convet-data-for-ai.gs`: スプレッドシートを PDF エクスポート/Google ドキュメント化

## 権限(Scopes)

```text
https://www.googleapis.com/auth/script.external_request
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/script.container.ui
https://www.googleapis.com/auth/userinfo.email
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/documents
https://www.googleapis.com/auth/forms
https://www.googleapis.com/auth/drive.readonly
https://www.googleapis.com/auth/script.scriptapp
https://www.googleapis.com/auth/script.send_mail
https://www.googleapis.com/auth/gmail.send
https://www.googleapis.com/auth/calendar
https://www.googleapis.com/auth/script.projects
https://www.googleapis.com/auth/documents.currentonly
```

## 事前準備(共通設定)

- **スクリプト プロパティ** に機密情報を保存
  - `SLACK_BOT_TOKEN`: Slack Bot Token (xoxb-...)
  - 任意: `YAHOO_CLIENT_ID` (形態素解析 API のアプリ ID) 例: `dj00aiZp...`
- **シート構成(存在が前提のタブ名)**
  - `フォルダ作成`, `キャンペーン一覧`, `スプシのスケ→teamXxxx一覧`, `チャンネル設定`, `text_input`, `text_output`, `作業用`, `INPUT`, `OUTPUT`
- **ID/URL の置換**
  - ファイル中の `SPREADSHEET_ID`, `folderId`, Slack Incoming Webhook URL 等のプレースホルダを自環境の値に置換
- **トリガー** は各スクリプトの「トリガー」節を参照して設定

---

## 共通ユーティリティ `google-spreadsheets/utils.gs`

- **UI**
  - なし(他スクリプトから関数/定数を使用)
- **ロジック**
  - シート参照: `spreadSheet`, `makeFolderSheet`, `campaignSheet`, `scheduleParseSheet`, `slackSheet`
  - Slack 送信: `sendSlack(text, team)` チャンネル種別 `pm|ope|kol|teamXxxx` で Webhook/表示名を切替
  - 日付/営業日/祝日: `isHoliday`, `getNextBusinessDay`, `getPreviousBusinessDay`
  - 列ヘッダー取得: `getHeaders(sheet, headersRow)` → ヘッダー名 → 列番号(
  - 形態素/集計補助: 判定/配列整形/降順ソート/グラフ作成 など
  - Slack Bot Token を `SLACK_BOT_TOKEN` から取得
- **依存**
  - 他多くのスクリプトが本ファイルの定数/関数に依存

---

## スクリプト別 仕様

### `api/tiktok.gs`

- **UI**
  - 手動実行関数: `tiktokAPI()`
- **ロジック**
  - アクティブシートの `H6` にある TikTok URL を oEmbed API で解決
  - サムネイル画像を (列 2, 行 13) に挿入(サイズ: 300x500)
  - アカウント名を `H4`、タイトルを `G13` に書き込み
- **入出力**
  - 入力: `H6`(URL)
  - 出力: 画像挿入/セル書き込み
- **トリガー**: なし(手動)
- **依存**: `UrlFetchApp`, `SpreadsheetApp`

### `google-documents/import-calendar.gs`

- **UI**
  - ドキュメント `onOpen` メニュー: `テンプレ追加 > ver1 / ver2 / ver3`
- **ロジック**
  - `insert_template_ver1|2|3`: タイトル/箇条書き/日付/カレンダー予定の挿入
  - `insert_template_ver3` は `CalendarApp` の当日予定(終日・昼休憩除外)を挿入
- **入出力**
  - 入力: 当日のカレンダー予定
  - 出力: ドキュメント本文への見出し/箇条書き/日付
- **トリガー**: ドキュメントを開いたとき(`onOpen`)
- **依存**: `DocumentApp`, `CalendarApp`

### `google-drive/create-folder.gs`

- **UI**
  - 手動実行関数: `search_clientFolder_ver1`, `make_clientFolder_ver1`, `brandAnalysis_makeFolder_duplicateSheet_ver2`
- **ロジック**
  - 会社/ブランド/商品階層のフォルダ自動作成と URL 出力
  - 既存検索(完全一致/部分一致)・ブランド分析テンプレの複製・作業用フォルダ作成
- **入出力**
  - 入力: `フォルダ作成` シート各セル
  - 出力: 作成したフォルダ URL をシートへ書き込み
- **トリガー**: なし(手動)
- **依存**
  - `utils.gs` の `makeFolderSheet`, `clientFolder`, `brandAnalysis_templateSheet`, `brandAnalysis_Folder`

### `google-forms/answer-limit.gs`

- **UI**: なし
- **ロジック**: 回答数が上限に達したら `form.setAcceptingResponses(false)`
- **トリガー**: 時間主導/手動(推奨: 定期実行)
- **依存**: `FormApp`

### `google-forms/default-value-1-3.gs`

- **UI**: なし
- **ロジック**
  - フォーム送信データからプレフィル URL を生成し Slack/シートへ出力する各バージョン
  - 設問タイプ(TEXT/PARAGRAPH/MULTIPLE_CHOICE/DATE/CHECKBOX)ごとに型変換して `createResponse`
- **前提/依存**
  - グローバル変数 `test_form`, `form_outputSheet`, `form_inputSheet` 等が同一プロジェクトで定義済みであること
  - `UrlFetchApp`(Slack Webhook 使用箇所)
- **トリガー**: フォーム送信時(必要関数をトリガーに割当)

### `google-forms/google-forms-to-slack/ver-1.gs`

- **UI**: なし
- **ロジック**: 送信内容を整形して Incoming Webhook に POST
- **入出力**: 入力=フォーム回答、出力=Slack 投稿
- **トリガー**: フォーム送信時 `autoSlack(e)`
- **注意**: Webhook URL はプロパティに移すことを推奨

### `google-forms/google-forms-to-slack/ver-2.gs`

- **UI**: なし
- **ロジック**
  - 回答の内容(媒体/キーワード)で送信先・メンション・ヘッダ文面を分岐(`getSendText`)
- **トリガー**: フォーム送信時 `autoSlack(e)`
- **依存**: 複数の Incoming Webhook URL/メンション文字列(定数)

### `google-spreadsheets/convert-schedules.gs`

- **UI**: なし
- **ロジック**
  - `スプシのスケ→teamXxxx一覧` のスケジュールから、`キャンペーン一覧` 指定行へ期日やフラグ/テキストを自動反映
  - `TeamXxxxFlagRange` を用いたフラグ集約、`getNextBusinessDay` 等で営業日計算
- **入出力**
  - 入力: スケジュール表, `メニュー`
  - 出力: 期日セルの日時・関連列のテキスト(「ー」埋め含む)
- **依存**: `utils.gs`(シート参照/クラス/日付関数)

### `google-spreadsheets/plot-calendar.gs`

- **UI**: なし
- **ロジック**: `INPUT` の範囲指定と開始日から `OUTPUT` に 4 ヶ月のカレンダーとタスクを描画(祝日/週末配慮)
- **入出力**: 入力=`INPUT`、出力=`OUTPUT`
- **依存**: `isJpHoliday`(ファイル内), `SpreadsheetApp`

### `google-spreadsheets/remind.gs`

- **UI**: なし
- **ロジック**
  - `キャンペーン一覧` を走査し、PM/OPE/KOL ごとの当日リマインド(予約公開/募集開始/納品/応募者リスト等)を組み立てて Slack 送信
  - 営業日/祝日判定で週末・祝日はスキップ
- **トリガー**: 時間主導(毎朝)
- **依存**: `utils.gs`(sendSlack/営業日関数), `campaignSheet`

### `google-spreadsheets/get_list.gs`

- **UI**: なし
- **ロジック**
  - 来月の薬事チェック案件を抽出し件数/合計人数を集計、Slack 投稿用テキストを作成
- **トリガー**: 時間主導(月次/週次)
- **依存**: `campaignSheet`, `getHeaders`(utils)

### `google-spreadsheets/get-holiday.gs`

- **UI**: なし
- **ロジック**
  - Google 公式祝日カレンダーから期間内イベントを取得し、`作業用` シートへ「日付/祝日名」または CSV 風日付配列を出力
- **トリガー**: 手動/必要に応じて
- **依存**: `CalendarApp`

### `google-spreadsheets/insert-updated-at.gs`

- **UI**: なし
- **ロジック**: 対象シートの更新時に、該当行の「更新日」列へ当日を記入(ヘッダー行は対象外)
- **トリガー**: スプレッドシートの `変更時(onEdit)`

### `google-spreadsheets/detect-header-change.gs`

- **UI**: なし
- **ロジック**: ヘッダー名の変更を検知し、差分の警告メッセージを Slack に送信/最新ヘッダ一覧の出力補助
- **トリガー**: 手動/時間主導
- **依存**: `campaignSheet`, `sendSlack`

### `google-spreadsheets/create-slack-channel.gs`

- **UI**: なし
- **ロジック**
  - チャンネル作成 → メンバー招待 → ブックマーク登録 →Canvas 作成 → 通知の一連を API で実行
- **トリガー**: 手動(一括実行は `mainSlackChannel()`)
- **依存/前提**
  - `token`: `SLACK_BOT_TOKEN` をプロパティに設定
  - `slackSheet`, `campaignSheet` の必要セルに値を用意

### `google-spreadsheets/prevent-archive.gs`

- **UI**: なし
- **ロジック**: 指定チャンネルに定期的なメッセージを送信(アーカイブ防止)
- **トリガー**: 時間主導(週 1/隔週など)

### `google-spreadsheets/morphological-analysis.gs`

- **UI**: なし
- **ロジック**
  - `text_input` のキャプションを Yahoo 形態素解析 API に投げ、語彙/品詞を集計し `text_output` に出力
  - 入力件数が多い場合はガード(150 件以下を推奨)
- **依存/前提**
  - `YAHOO_CLIENT_ID` をプロパティに保存して取得する実装に差し替え推奨
  - `utils.gs` の補助関数を使用

### `google-spreadsheets/extract-keyword.gs`

- **UI**: なし
- **ロジック**
  - 正規表現で PR/AF/ORG/ALL の該当語を集計し、出力タブにランキング/円グラフを生成
- **依存**: `utils.gs` の正規表現/整形/グラフ関数

### `google-spreadsheets/check-for-updates-with-redash-api.gs`

### `google-spreadsheets/detect-updating-sheets.gs`

- **UI**: なし
- **ロジック**
  - `SPREADSHEET_ID` と `SHEET_NAME` のセル値差分を前回スナップショットと比較し、変更点を LCS で抽出して Slack 通知
  - HTML を `<br>` 区切りで行比較する解析ユーティリティを同梱
- **トリガー**: 時間主導(例: 毎日 09:00)
- **依存**: `PropertiesService`, `UrlFetchApp`

### `google-spreadsheets/convet-data-for-ai.gs`

- **UI**: なし
- **ロジック**
  - スプレッドシートを PDF としてエクスポートし、指定フォルダへ保存/リンクをシートへ書き出し
  - シート内容を単純に Google ドキュメント化するユーティリティ
- **前提**: `spreadsheetId`, `folderId`, `sheetName` を自環境に合わせて設定

---

## 推奨トリガー設定一覧

- **フォーム送信時**: `google-forms/google-forms-to-slack/ver-*.gs` の `autoSlack(e)`、`default-value-1-3.gs` の送信ハンドラ
- **ドキュメントを開いたとき**: `google-documents/import-calendar.gs` の `onOpen`
- **スプレッドシート編集時**: `insert-updated-at.gs` の `insertLastModified`
- **時間主導(毎朝 09:00)**: `remind.gs` のリマインド、`check-for-updates-*.gs` の更新検知、`prevent-archive.gs`
- **必要時手動**: `create-folder.gs`/`create-slack-channel.gs`/`plot-calendar.gs` など

---

## セキュリティ/運用上の注意

- **Webhook URL / API トークンはハードコードしない**: スクリプト プロパティに保存し、コードでは取得して使用
- **最小権限**: 不要な Scope は付与しない(本リポジトリの `appsscript.json` を参考に用途に応じて絞る)
- **コンテナ分離**: 可能であれば フォーム/ドキュメント/スプレッドシート でプロジェクトを分け、グローバル関数名の衝突や `getActive*` の対象誤りを防止
- **ID/タブ名の前提**: `utils.gs` で参照するタブ名・ID を環境に合わせて調整

---

## 動作確認チェックリスト

- **シート/タブ**: 必須タブが存在し、想定列ヘッダーが一致している
- **プロパティ**: `SLACK_BOT_TOKEN` 他のキーを登録済み
- **トリガー**: 想定の時間/イベントに設定済み
- **テスト**: 手動実行で 1〜2 件のデータを用いて Slack 送信/書き込み/作成処理を確認

---

## ライセンス

- 本リポジトリのコードは学習/業務ユースのサンプルです。利用にあたっては API 利用規約/社内規程に従ってください。

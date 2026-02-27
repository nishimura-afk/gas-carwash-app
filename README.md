# 洗車機管理システム

ガソリンスタンドの洗車機に関する設備管理を行うWebアプリ。

## 何ができるか

- 洗車機の部品（レールホイール、布ブラシ、機体本体など）の交換時期を管理する
- 交換スケジュールを自動計算し、カレンダーに登録する
- 月次データの更新と進捗管理
- 補助金対象店舗のアラート表示
- メール下書きの自動作成

## 技術

- Google Apps Script（clasp で管理）
- スプレッドシートをデータベースとして使用
- Gmail API・Drive API・Calendar 連携
- Webアプリとしてデプロイ

## ファイル構成

- `0_Config.js` … 設定（スプレッドシートID、交換基準、補助金対象店舗）
- `1_Setup.js` … 初期セットアップ
- `2_DataService.js` … データ読み書き
- `3_DataUpdateService.js` … 月次データ更新
- `4_CalendarService.js` … カレンダー連携
- `Code.js` … Webアプリのメイン処理（doGet、ルーター）
- `index.html` … トップページ
- `dashboard.html` … ダッシュボード画面
- `schedule.html` … 交換スケジュール画面

## 開発方法

```
clasp push    # GASにコードを反映
clasp pull    # GASからコードを取得
```

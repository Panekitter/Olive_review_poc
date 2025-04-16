# Olive_review_poc

このプロジェクトはGoogleスプレッドシートに記載されたレビュー対象ファイル（URL）をGPT-4oでチェックし、翻訳レビューと修正を行うPoCです。

## 🔧 必要なSecrets

| Name | 説明 |
|------|------|
| `OPENAI_API_KEY` | OpenAIのAPIキー |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | サービスアカウントのJSON内容（1行にまとめる） |
| `MASTER_SPREADSHEET_URL` | URL一覧を記載したスプレッドシートのURL |

## ▶ 実行方法

1. Secretsを登録
2. GitHub Actions > Run workflow をクリック
3. 各URL先のスプレッドシートの「Task」シートが処理され、C〜E列にレビュー結果が記入されます。

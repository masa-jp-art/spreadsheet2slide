# Google Slides 自動作成ツール

## 概要
- このツールは、Googleスプレッドシートのデータを基にGoogleスライドを自動生成するPythonスクリプトです。スプレッドシートの各行からスライドを作成し、タイトル、サブタイトル、本文を適切なフォーマットで配置します。
- AIアートグランプリの最終選考会で展示する作品制作のために作成しました
- コード作成は、GeminiとClaude3.5Sonnet(new)の補助を受けています
- README文はClaude3.5Sonnet(new)が生成したものを修正しています
- 実際に使用したコードをapp.pyとして、Claude3.5Sonnet(new)がリファクタリングしたものをapp-rf.pyとしています

## 主な機能
- スプレッドシートからデータを読み込み
- スライドの自動生成
- カスタマイズされたテキストボックスの配置
- エラーハンドリングとログ記録

## 必要条件
- Python 3.7以上
- Google Colabの実行環境
- 必要なPythonパッケージ:
  - gspread
  - google-auth-oauthlib
  - google-api-python-client

## インストール方法
```bash
pip install gspread google-auth-oauthlib google-api-python-client
```

## 使用方法

### 1. 設定
以下の情報を`main()`関数内で設定します：
```python
SPREADSHEET_ID = '<GoogleスプレッドシートのID>'
SHEET_NAME = '<シート名>'
PRESENTATION_ID = '<GoogleスライドのID>'
```

### 2. スプレッドシートの準備
- 1行目: ヘッダー行
- 2行目以降: 各スライドのコンテンツ
  - A列: タイトル
  - B列: サブタイトル
  - C列: 本文

### 3. 実行
```python
python main.py
```

## テキストボックスの仕様

### タイトル用テキストボックス
- サイズ: 600pt × 50pt
- 位置: (30, 30)
- フォント: BIZ UDPMincho, 30pt
- スタイル: 太字, 左揃え

### サブタイトル用テキストボックス
- サイズ: 400pt × 240pt
- 位置: (30, 80)
- フォント: BIZ UDPGothic, 12pt
- スタイル: 通常, 左揃え

### 本文用テキストボックス
- サイズ: 400pt × 80pt
- 位置: (30, 320)
- フォント: BIZ UDPMincho, 14pt
- スタイル: 太字, 左揃え

## エラーハンドリング
- 空の行は自動的にスキップ
- 処理エラー時は該当行をスキップして続行
- 詳細なエラーログを出力

## 注意事項
- Google認証が必要です
- API制限に注意してください
- 大量のデータ処理時は適切な待機時間を設定してください

## トラブルシューティング
1. 認証エラーの場合
   - Google認証情報を確認
   - 必要な権限が付与されているか確認

2. API制限エラーの場合
   - 待機時間（time.sleep）を調整
   - 一度に処理する量を制限

3. フォーマットエラーの場合
   - スプレッドシートのデータ形式を確認
   - 必須項目が空でないか確認

## 貢献
バグ報告や機能改善の提案は、IssuesまたはPull Requestsでお願いします。

## ライセンス
MITライセンス

# コード

## 実際に使用したもの

- https://github.com/masa-jp-art/spreadsheet2slide/blob/main/app.py

## Claude3.5Sonnet(new)がリファクタリングしたもの

- https://github.com/masa-jp-art/spreadsheet2slide/blob/main/app-rf.py

# 関連

- 第3回AIアートグランプリエントリー作品
- https://github.com/masa-jp-art/100-times-ai-heroes

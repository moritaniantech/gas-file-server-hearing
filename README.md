# ファイルサーバーヒアリング GAS

Google Apps Scriptを使用したファイルサーバーヒアリング用のスクリプトです。

## 概要

このプロジェクトは、MIXIのファイルサーバーヒアリング用のGoogle Apps Scriptです。
divフォルダとprojectフォルダの処理を行う2つのスクリプトが含まれています。

## ファイル構成

- `scripts/div.gs` - divフォルダ処理用スクリプト（共通メニュー含む）
- `scripts/project.gs` - projectフォルダ処理用スクリプト
- `appsscript.json` - Apps Scriptの設定ファイル

## セットアップ

### 1. Claspのインストール

```bash
npm install -g @google/clasp
```

### 2. Claspでログイン

```bash
clasp login
```

### 3. プロジェクトの作成/接続

新しいプロジェクトを作成する場合：
```bash
clasp create --type sheets --title "ファイルサーバーヒアリングGAS" --rootDir scripts
```

既存のプロジェクトに接続する場合：
```bash
clasp clone <SCRIPT_ID>
```

### 4. スクリプトのプッシュ

```bash
clasp push
```

## 使用方法

1. Googleスプレッドシートを開く
2. 「GAS実行メニュー」から実行したい処理を選択
   - div フォルダ処理
     - 回答シート作成
     - 回答シートマージ
   - project フォルダ処理
     - 回答シート作成
     - 回答シートマージ

## 注意事項

- フォルダIDの設定が必要です（各スクリプト内の定数を確認してください）
- 実行前に必要なシートが存在することを確認してください


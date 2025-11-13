# ファイルを手動で追加・更新する手順

claspが未追跡ファイルを自動的に追加しない場合、以下の手順でApps Scriptエディタから手動で追加・更新できます。

## 新規ファイルを追加する場合

1. **Apps Scriptエディタを開く**
   - https://script.google.com/d/1vUyb3TrBm57302puLaxFaDFVofq3kmF9Lo5wNBdmcs3LUJjh-5O5ygiO/edit

2. **divファイルを追加**
   - 左側の「+」ボタンをクリック
   - 「スクリプト」を選択
   - ファイル名を「div」に変更
   - ローカルの`scripts/div.js`の内容をコピー＆ペースト

3. **projectファイルを追加**
   - 左側の「+」ボタンをクリック
   - 「スクリプト」を選択
   - ファイル名を「project」に変更
   - ローカルの`scripts/project.js`の内容をコピー＆ペースト

4. **保存**
   - すべてのファイルを保存（Ctrl+S / Cmd+S）

## 既存ファイルを更新する場合

1. **Apps Scriptエディタを開く**
   - https://script.google.com/d/1vUyb3TrBm57302puLaxFaDFVofq3kmF9Lo5wNBdmcs3LUJjh-5O5ygiO/edit

2. **projectファイルを更新**
   - 左側のファイル一覧で「project」を選択
   - ローカルの`scripts/project.js`の内容をコピーして貼り付け（既存の内容を全て置き換え）
   - 保存（Ctrl+S / Cmd+S）

3. **divファイルを更新**
   - 左側のファイル一覧で「div」を選択
   - ローカルの`scripts/div.js`の内容をコピーして貼り付け（既存の内容を全て置き換え）
   - 保存（Ctrl+S / Cmd+S）

## 手動更新後の確認

手動で更新した後、以下のコマンドでローカルに反映されているか確認します：

```bash
cd "/Users/ryosuke.morita.ts/Documents/Dev/GAS - ファイルサーバーヒアリング"
clasp pull
```

これにより、Google Apps Script側の最新の状態がローカルに取得されます。


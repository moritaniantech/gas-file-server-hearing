# 既存のSpreadsheetに接続する方法

## Script IDの取得方法

1. 既存のGoogleスプレッドシートを開く
2. **拡張機能** > **Apps Script** をクリック
3. プロジェクト設定（歯車アイコン）をクリック
4. **スクリプトID** をコピー

または、Apps ScriptエディタのURLから取得：
```
https://script.google.com/d/{SCRIPT_ID}/edit
```

## 接続方法

### 方法1: clasp cloneを使用（推奨）

```bash
# 現在の.clasp.jsonを削除またはバックアップ
mv .clasp.json .clasp.json.backup

# 既存プロジェクトをクローン
clasp clone <SCRIPT_ID> --rootDir scripts
```

### 方法2: .clasp.jsonを直接更新

Script IDとSpreadsheet IDが分かっている場合：

```bash
# .clasp.jsonを編集
# scriptId: 既存のScript ID
# parentId: 既存のSpreadsheet ID（オプション）
```

その後、`clasp pull`で既存のコードを取得できます。


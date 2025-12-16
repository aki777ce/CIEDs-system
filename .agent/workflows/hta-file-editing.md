---
description: HTAファイル編集時のエンコーディングルール（文字化け防止）
---

# HTAファイル編集ワークフロー

このプロジェクトのHTAファイル（`CIED_App.hta`）を編集する際は、以下のルールを**必ず**遵守すること。

## 重要なルール

### 1. エンコーディング
- HTAファイルは **Shift_JIS（Windows標準）** エンコーディングで保存すること
- UTF-8で保存すると日本語が文字化けする

### 2. ファイル保存後の変換手順
HTAファイルを `write_to_file` や `replace_file_content` で編集した後、**必ず**以下のコマンドを実行してエンコーディングを変換する：

```powershell
// turbo
$content = Get-Content "c:\Users\aki77\Antigravity projects\CIEDs-system\CIED_App.hta" -Encoding UTF8 -Raw; $content | Out-File -FilePath "c:\Users\aki77\Antigravity projects\CIEDs-system\CIED_App.hta" -Encoding Default
```

### 3. VBSファイルも同様
`CreateDatabase.vbs` などのVBScriptファイルも同様にShift_JISで保存する必要がある：

```powershell
// turbo
$content = Get-Content "c:\Users\aki77\Antigravity projects\CIEDs-system\CreateDatabase.vbs" -Encoding UTF8 -Raw; $content | Out-File -FilePath "c:\Users\aki77\Antigravity projects\CIEDs-system\CreateDatabase.vbs" -Encoding Default
```

## 対象ファイル一覧
| ファイル名 | エンコーディング | 備考 |
|-----------|------------------|------|
| `CIED_App.hta` | Shift_JIS | メインアプリケーション |
| `CreateDatabase.vbs` | Shift_JIS | データベース初期化スクリプト |
| `CIED_DB.xlsx` | - | Excelファイル（バイナリ） |

## チェックリスト
ファイル編集後、以下を確認すること：
- [ ] エンコーディング変換コマンドを実行した
- [ ] HTAファイルを開いて日本語が正しく表示されることを確認した
- [ ] アプリケーションを起動して動作確認した

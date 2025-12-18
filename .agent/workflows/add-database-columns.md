---
description: データベースのカラム（項目）を追加・変更する際の手順
---

# データベースカラム追加・変更ワークフロー

データベース（Excel）に新しい項目を追加し、アプリケーション（HTA）に反映させるための標準手順です。

## 1. データベース定義の更新 (CreateDatabase.vbs)

`CreateDatabase.vbs` スクリプトを編集し、新カラムを追加します。

1.  **カラム名の命名則**:
    *   **必ず英語（CamelCase）** で命名する（例: `PatientID`, `NameKana`, `RegistrationDate`）。
    *   *理由*: JScript内でのオブジェクトプロパティアクセス (`p.Name`) を確実にするため、およびマルチバイト文字によるスクリプトエラーを防ぐため。
2.  **ヘッダー配列の更新**: `headersPatients` などの配列に項目を追加。
3.  **サンプルデータの更新**:
    *   新カラムに合わせて `objSheet.Cells(2, n).Value` を追加。
    *   **電話番号やID** は、Excelが自動で数値や日付に変換するのを防ぐため、先頭にシングルクォーテーションを付けて文字列として扱う（例: `objSheet.Cells(2, 8).Value = "'03-1234-5678"`）。
4.  **保存エンコーディング**: 必ず **Shift_JIS (Encoding Default)** で保存すること。

## 2. データベースの再生成

1.  既存の Excel インスタンスが残っているとエラーになるため、タスクマネージャーで閉じるか、以下のコマンドで強制終了させてから実行する。
    ```powershell
    taskkill /F /IM excel.exe
    cscript CreateDatabase.vbs
    ```

## 3. アプリケーションの更新 (CIED_App.hta)

### A. データ表示ロジックの修正
1.  **詳細画面 (`selectPatient`)**:
    *   `detail-value` クラスの項目を追加。
    *   `p["EnglishColumnName"]` または `(p["Name"] || "-")` の形式で値を埋め込む。
2.  **年齢計算**: 生年月日の場合は `calculateAge(p["BirthDate"])` を活用する。
3.  **日付フォーマット**: `formatDate(key, val)` 関数が、追加したカラム名（key）を日付として認識するようにリストを更新する。

### B. 保存ロジックの修正 (`savePatient`)
1.  **入力取得**: UIから新項目の `value` を取得。
2.  **Excel書き込み**: `sheet.Cells(lastRow, n).Value` で順番に書き込む。電話番号などはここでも `"'"` を付与することを検討する。

### C. 新規登録モーダル
1.  `showNewPatientModal` 内の HTML を更新し、`input` 要素を追加する。

## 4. 文字化け防止策
ファイルを編集する際は、以下のPowerShellコマンドを使用してエンコーディングを維持すること。
```powershell
$content = Get-Content "path\to\file" -Raw; $content | Out-File -FilePath "path\to\file" -Encoding Default
```

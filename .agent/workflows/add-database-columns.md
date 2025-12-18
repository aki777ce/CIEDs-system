---
description: データベースのカラム（項目）を追加・変更する際の手順
---

# データベースカラム追加・変更ワークフロー

データベース（Excel）に新しい項目を追加し、アプリケーション（HTA）に反映させるための標準手順です。

## 1. データベース定義の更新 (CreateDatabase.vbs)

`CreateDatabase.vbs` スクリプトを開き、該当するテーブルのヘッダー定義配列を修正します。

1.  ファイルを開く: `c:\Users\aki77\Antigravity projects\CIEDs-system\CreateDatabase.vbs`
2.  対象の配列を見つける:
    *   患者基本情報: `headersPatient`
    *   デバイス情報: `headersDevice` (もしあれば)
    *   計測値: `headersMeasurement`
3.  配列に新しいカラム名（文字列）を追加する。
    *   **注意**: カラム名は基本的に英語/ローマ字推奨（例: `PatientID`, `Name`, `BirthDate`）。日本語もExcel上は動作するが、SQLクエリでの扱いやすさを考慮する。

```vbscript
' 例: 生年月日を追加する場合
headersPatient = Array("PatientID", "Name", "Kana", "BirthDate", "Sex", ...)
```

4.  **重要**: ファイルを保存する際は **Shift_JIS** エンコーディングを維持すること。

## 2. データベースの再生成

**注意**: この操作を行うと、既存のデータはすべて初期化（削除）されます。開発段階でのみ推奨されます。

1.  ターミナルで `CreateDatabase.vbs` を実行する。
    ```powershell
    cscript CreateDatabase.vbs
    ```
2.  `CIED_DB.xlsx` が新しく生成されたことを確認する。

## 3. アプリケーションの更新 (CIED_App.hta)

`CIED_App.hta` を開き、以下の箇所を修正する。（**Shift_JIS** 保存必須）

### A. データ読み込み (Load)
`loadPatients` 関数または詳細表示を行う関数を確認する。
*   `SELECT *` を使用している場合、SQLの変更は不要。
*   明示的にカラムを指定している場合は、SQL文に新カラムを追加する。

### B. 一覧表示 (List View)
一覧に表示したい項目であれば、`renderPatientList` 関数内のHTML生成部分に新しいセル `<td>` を追加する。
```javascript
html += "<td class='col-birth'>" + (rs.Fields("BirthDate").Value || "") + "</td>";
```

### C. 詳細画面・入力フォーム (Detail/Edit View)
詳細画面 (`showPatientDetail` など) のフォームHTMLを修正し、新しい入力フィールドを追加する。
```html
<!-- 例 -->
<div class="form-group">
    <label>生年月日</label>
    <input type="text" id="input-BirthDate" value="">
</div>
```
既存データの表示ロジックも追加する。
```javascript
document.getElementById("input-BirthDate").value = rs.Fields("BirthDate").Value || "";
```

### D. 保存処理 (Save)
データの保存関数（`savePatient` など）を更新し、フォームの値をデータベースに書き込む処理を追加する。

**ADODB RecordSetを使用している場合:**
```javascript
rs.Fields("BirthDate").Value = document.getElementById("input-BirthDate").value;
rs.Update();
```

**SQL UPDATE文を使用している場合:**
```javascript
var sql = "UPDATE T_Patients SET BirthDate='" + val + "' WHERE ...";
```

## 4. 動作確認

1.  アプリを起動する。
2.  新規登録または編集を行い、新しい項目が正しく保存されるか確認する。
3.  アプリを再起動し、保存した値が正しく読み込まれるか確認する。

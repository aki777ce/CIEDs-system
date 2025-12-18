# CIEDs Patient Management System - 開発アーキテクチャドキュメント

## 1. プロジェクト概要
本システムは、インターネット接続のない閉鎖的な医療環境（電子カルテ端末等）において、共有フォルダ上に配置するだけで動作する「インストール不要」「完全オフライン」の患者管理データベースです。

## 2. 制約条件と技術選定
度重なる検証の結果、以下の構成を**最終決定**として採用しています。今後の開発で安易に変更しないこと。

*   **プラットフォーム**: Windows HTA (HTML Application)
*   **使用言語**: **JScript (JavaScript)** のみ
    *   *理由*: VBScriptは環境によって無効化されている場合があるため使用不可。
*   **バックエンド**: Excelファイル (`CIED_DB.xlsx`)
*   **データアクセス方式**: **Excel Automation (Excel.Application)**
    *   *理由*: ADODB / ACE.OLEDB ドライバは、Officeのバージョン (2016/2019/365) やビット数 (32bit/64bit) の違いにより接続できないトラブルが多発するため廃止。
    *   *採用理由*: Excelがインストールされている端末であればバージョン問わず確実に動作するため。

## 3. システムアーキテクチャ

### 3.1 ディレクトリ構成
```text
/ProjectRoot
  │
  ├─ CIED_App.hta       # メインアプリケーション (UI + Logic)
  ├─ CIED_DB.xlsx       # データベース (T_Patients, T_DeviceConfig, etc.)
  └─ CIED_DB.lock       # 排他制御用ロックファイル (書き込み時のみ生成)
```

### 3.2 データ読み込み (Read)
*   **起動時一括ロード**:
    *   Excel Automationの起動は遅いため、アプリ起動時に一度だけExcelを裏で開き、全データをメモリ（JavaScriptオブジェクト配列）に読み込みます。
    *   以降のリスト表示、検索、詳細表示はメモリ上のデータを使用するため**高速**です。
*   **ReadOnlyモード**:
    *   Excelを開く際は `ReadOnly=True` を指定し、他ユーザーのロックを誘発しないようにします。

### 3.3 データ書き込み (Write)
*   **排他制御 (Locking)**:
    *   保存処理の開始時、`CIED_DB.lock` ファイルの存在を確認・生成します。
    *   ロックファイルが存在する場合は「他ユーザーが編集中」として処理を中断し、ユーザーにリトライを促します。
*   **書き込みプロセス**:
    1.  ロック取得
    2.  Excelを通常モード（Read/Write）で開く
    3.  最終行に追記
    4.  保存して閉じる
    5.  ロック解除
    6.  全データを再ロード（リフレッシュ）

## 4. データベース定義
各シートの1行目をヘッダーとして扱います。

*   **T_Patients**: 患者基本情報
    *   `PatientID`, `Name`, `NameKana`, `Sex`, `BirthDate`, `PostalCode`, `Address`, `Phone`, `MRICompatible`, `RMS_Enabled`, `RMS_TransmitterType`, `RegistrationDate`
*   **T_DeviceConfig**: デバイス植込み履歴
    *   `ConfigID`, `PatientID`, `手術日`, `本体メーカー`, `本体型番`, `本体Serial`, `ステータス` (注: 内部項目は順次英語化検討)
*   **T_Settings**: 設定履歴
    *   `SettingID`, `PatientID`, `設定日`, `Mode`, `LowerRate`, `UpperRate`, `AV_Delay`
*   **T_Measurements**: 測定履歴
    *   `MeasureID`, `PatientID`, `計測日`, `Sensing_A`, `Sensing_V`, `Impedance_A`, `Impedance_RV`

## 5. 開発時の重要ルール (Do's & Don'ts)

*   **DO (やるべきこと)**:
    *   **Shift_JIS保存**: HTAファイルは必ず **Shift_JIS (ANSI)** エンコーディングで保存してください。UTF-8では文字化けし、スクリプトエラーの原因になります。
        *   PowerShellコマンド: `$content | Out-File -FilePath "..." -Encoding Default`
    *   **JScriptの使用**: 新機能追加時は必ず `<script language="JavaScript">` ブロック内に記述してください。

*   **DON'T (やってはいけないこと)**:
    *   **VBScriptの使用**: 互換性担保のため使用禁止。
    *   **ADODB接続の復活**: 環境依存エラーの原因となるため禁止。
    *   **外部ライブラリ依存**: jQueryやBootstrapなどはCDN経由で使えないため、すべてVanilla JS / CSSで実装すること。

## 6. 未実装機能ロードマップ
*   [ ] デバイス情報の新規追加・編集モーダル
*   [ ] 設定情報の新規追加・編集モーダル
*   [ ] 測定情報の新規追加・編集モーダル
*   [ ] CSVファイルインポート機能（測定データの取り込み）
*   [ ] PDF/画像ファイルの参照機能（フォルダパス紐付け）

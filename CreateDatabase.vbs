' ============================================
' CIEDs Database Initialization Script
' CreateDatabase.vbs
' ============================================
' このスクリプトはCIED_DB.xlsxを初期生成します。
' 実行方法: ダブルクリック または cscript CreateDatabase.vbs
' ============================================

Option Explicit

Dim objExcel, objWorkbook, objSheet
Dim strPath, strDbPath
Dim arrSheets, arrHeaders
Dim i, j

' スクリプトのあるフォルダを取得
strPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
strDbPath = strPath & "CIED_DB.xlsx"

' 既存ファイルの確認
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strDbPath) Then
    If MsgBox("CIED_DB.xlsx は既に存在します。上書きしますか？", vbYesNo + vbQuestion, "確認") = vbNo Then
        WScript.Echo "処理を中止しました。"
        WScript.Quit
    End If
    objFSO.DeleteFile strDbPath, True
End If

' Excelアプリケーション起動
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Excelを起動できませんでした。" & vbCrLf & "Microsoft Excelがインストールされているか確認してください。", vbCritical, "エラー"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

' 新規ワークブック作成
Set objWorkbook = objExcel.Workbooks.Add

' シート定義
arrSheets = Array("T_Patients", "T_DeviceConfig", "T_Settings", "T_Measurements", "T_Documents")

' 各シートのヘッダー定義
Dim headersPatients, headersDeviceConfig, headersSettings, headersMeasurements, headersDocuments

headersPatients = Array("PatientID", "カルテ番号", "氏名", "生年月日", "性別", "原疾患", "住所", "電話番号")

headersDeviceConfig = Array("ConfigID", "PatientID", "手術日", "本体メーカー", "本体型番", "本体Serial", "リード構成", "ステータス")

headersSettings = Array("SettingID", "PatientID", "設定日", "Mode", "LowerRate", "UpperRate", "Output_A", "Output_V", "Sensitivity_A", "Sensitivity_V", "AV_Delay", "ZoneSettings", "Details")

headersMeasurements = Array("MeasureID", "PatientID", "計測日", "区分", "Battery", "Impedance_A", "Impedance_RV", "Impedance_LV", "Threshold_A", "Threshold_RV", "Threshold_LV", "Sensing_A", "Sensing_V", "Burden_ATAF", "Burden_VT")

headersDocuments = Array("DocID", "PatientID", "登録日", "カテゴリ", "FilePath")

' 必要なシート数を確保
Do While objWorkbook.Sheets.Count < UBound(arrSheets) + 1
    objWorkbook.Sheets.Add , objWorkbook.Sheets(objWorkbook.Sheets.Count)
Loop

' 不要なシートを削除（3枚がデフォルトで作成される場合）
Do While objWorkbook.Sheets.Count > UBound(arrSheets) + 1
    objWorkbook.Sheets(objWorkbook.Sheets.Count).Delete
Loop

' 各シートの設定
For i = 0 To UBound(arrSheets)
    Set objSheet = objWorkbook.Sheets(i + 1)
    objSheet.Name = arrSheets(i)
    
    ' ヘッダー設定
    Select Case arrSheets(i)
        Case "T_Patients"
            For j = 0 To UBound(headersPatients)
                objSheet.Cells(1, j + 1).Value = headersPatients(j)
            Next
        Case "T_DeviceConfig"
            For j = 0 To UBound(headersDeviceConfig)
                objSheet.Cells(1, j + 1).Value = headersDeviceConfig(j)
            Next
        Case "T_Settings"
            For j = 0 To UBound(headersSettings)
                objSheet.Cells(1, j + 1).Value = headersSettings(j)
            Next
        Case "T_Measurements"
            For j = 0 To UBound(headersMeasurements)
                objSheet.Cells(1, j + 1).Value = headersMeasurements(j)
            Next
        Case "T_Documents"
            For j = 0 To UBound(headersDocuments)
                objSheet.Cells(1, j + 1).Value = headersDocuments(j)
            Next
    End Select
    
    ' ヘッダー行の書式設定
    objSheet.Rows(1).Font.Bold = True
    objSheet.Rows(1).Interior.Color = RGB(70, 130, 180) ' SteelBlue
    objSheet.Rows(1).Font.Color = RGB(255, 255, 255)    ' White
    
    ' 列幅の自動調整
    objSheet.Columns.AutoFit
Next

' サンプルデータの挿入（動作確認用）
Set objSheet = objWorkbook.Sheets("T_Patients")
objSheet.Cells(2, 1).Value = "P0001"           ' PatientID
objSheet.Cells(2, 2).Value = "12345678"        ' カルテ番号
objSheet.Cells(2, 3).Value = "山田 太郎"       ' 氏名
objSheet.Cells(2, 4).Value = "1950/01/15"      ' 生年月日
objSheet.Cells(2, 5).Value = "男"              ' 性別
objSheet.Cells(2, 6).Value = "完全房室ブロック" ' 原疾患
objSheet.Cells(2, 7).Value = "東京都千代田区1-1-1" ' 住所
objSheet.Cells(2, 8).Value = "03-1234-5678"    ' 電話番号

Set objSheet = objWorkbook.Sheets("T_DeviceConfig")
objSheet.Cells(2, 1).Value = "C0001"           ' ConfigID
objSheet.Cells(2, 2).Value = "P0001"           ' PatientID
objSheet.Cells(2, 3).Value = "2020/05/15"      ' 手術日
objSheet.Cells(2, 4).Value = "Medtronic"       ' 本体メーカー
objSheet.Cells(2, 5).Value = "Azure XT DR MRI" ' 本体型番
objSheet.Cells(2, 6).Value = "ABC123456"       ' 本体Serial
objSheet.Cells(2, 7).Value = "RA-RV"           ' リード構成
objSheet.Cells(2, 8).Value = "Active"          ' ステータス

Set objSheet = objWorkbook.Sheets("T_Settings")
objSheet.Cells(2, 1).Value = "S0001"           ' SettingID
objSheet.Cells(2, 2).Value = "P0001"           ' PatientID
objSheet.Cells(2, 3).Value = "2020/05/15"      ' 設定日
objSheet.Cells(2, 4).Value = "DDDR"            ' Mode
objSheet.Cells(2, 5).Value = 60                ' LowerRate
objSheet.Cells(2, 6).Value = 130               ' UpperRate
objSheet.Cells(2, 7).Value = "2.5V/0.4ms"      ' Output_A
objSheet.Cells(2, 8).Value = "2.5V/0.4ms"      ' Output_V
objSheet.Cells(2, 9).Value = "0.5mV"           ' Sensitivity_A
objSheet.Cells(2, 10).Value = "2.5mV"          ' Sensitivity_V
objSheet.Cells(2, 11).Value = 180              ' AV_Delay
objSheet.Cells(2, 12).Value = ""               ' ZoneSettings
objSheet.Cells(2, 13).Value = ""               ' Details

Set objSheet = objWorkbook.Sheets("T_Measurements")
objSheet.Cells(2, 1).Value = "M0001"           ' MeasureID
objSheet.Cells(2, 2).Value = "P0001"           ' PatientID
objSheet.Cells(2, 3).Value = "2024/01/10"      ' 計測日
objSheet.Cells(2, 4).Value = "外来"            ' 区分
objSheet.Cells(2, 5).Value = "3.0V"            ' Battery
objSheet.Cells(2, 6).Value = 450               ' Impedance_A
objSheet.Cells(2, 7).Value = 520               ' Impedance_RV
objSheet.Cells(2, 8).Value = ""                ' Impedance_LV
objSheet.Cells(2, 9).Value = "0.5V/0.4ms"      ' Threshold_A
objSheet.Cells(2, 10).Value = "0.75V/0.4ms"    ' Threshold_RV
objSheet.Cells(2, 11).Value = ""               ' Threshold_LV
objSheet.Cells(2, 12).Value = "2.5mV"          ' Sensing_A
objSheet.Cells(2, 13).Value = "8.0mV"          ' Sensing_V
objSheet.Cells(2, 14).Value = "0%"             ' Burden_ATAF
objSheet.Cells(2, 15).Value = "0%"             ' Burden_VT

' ファイル保存
objWorkbook.SaveAs strDbPath, 51  ' 51 = xlOpenXMLWorkbook (.xlsx)

' クリーンアップ
objWorkbook.Close False
objExcel.Quit

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing

MsgBox "CIED_DB.xlsx を正常に作成しました。" & vbCrLf & vbCrLf & _
       "場所: " & strDbPath & vbCrLf & vbCrLf & _
       "サンプルデータ（山田太郎）が登録されています。", vbInformation, "完了"

WScript.Quit

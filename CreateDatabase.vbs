' ============================================
' CIEDs Database Initialization Script
' CreateDatabase.vbs
' ============================================
Option Explicit

Dim objExcel, objWorkbook, objSheet
Dim strPath, strDbPath
Dim arrSheets
Dim i, j

strPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
strDbPath = strPath & "CIED_DB.xlsx"

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strDbPath) Then
    objFSO.DeleteFile strDbPath, True
End If

On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Excel Error"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objWorkbook = objExcel.Workbooks.Add
arrSheets = Array("T_Patients", "T_DeviceConfig", "T_Settings", "T_Measurements", "T_Documents")

Dim headersPatients, headersDeviceConfig, headersSettings, headersMeasurements, headersDocuments
headersPatients = Array("PatientID", "Name", "NameKana", "Sex", "BirthDate", "PostalCode", "Address", "Phone", "MRICompatible", "RMS_Enabled", "RMS_TransmitterType", "RegistrationDate")
headersDeviceConfig = Array("ConfigID", "PatientID", "手術日", "本体メーカー", "本体型番", "本体Serial", "リード構成", "ステータス")
headersSettings = Array("SettingID", "PatientID", "設定日", "Mode", "LowerRate", "UpperRate", "Output_A", "Output_V", "Sensitivity_A", "Sensitivity_V", "AV_Delay", "ZoneSettings", "Details")
headersMeasurements = Array("MeasureID", "PatientID", "計測日", "区分", "Battery", "Impedance_A", "Impedance_RV", "Impedance_LV", "Threshold_A", "Threshold_RV", "Threshold_LV", "Sensing_A", "Sensing_V", "Burden_ATAF", "Burden_VT")
headersDocuments = Array("DocID", "PatientID", "登録日", "カテゴリ", "FilePath")

Do While objWorkbook.Sheets.Count < UBound(arrSheets) + 1
    objWorkbook.Sheets.Add , objWorkbook.Sheets(objWorkbook.Sheets.Count)
Loop
Do While objWorkbook.Sheets.Count > UBound(arrSheets) + 1
    objWorkbook.Sheets(objWorkbook.Sheets.Count).Delete
Loop

For i = 0 To UBound(arrSheets)
    Set objSheet = objWorkbook.Sheets(i + 1)
    objSheet.Name = arrSheets(i)
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
    objSheet.Rows(1).Font.Bold = True
    objSheet.Columns.AutoFit
Next

Set objSheet = objWorkbook.Sheets("T_Patients")
objSheet.Cells(2, 1).Value = "P0001"
objSheet.Cells(2, 2).Value = "山田 太郎"
objSheet.Cells(2, 3).Value = "ヤマダ タロウ"
objSheet.Cells(2, 4).Value = "男"
objSheet.Cells(2, 5).Value = "1950/01/15"
objSheet.Cells(2, 6).Value = "100-0001"
objSheet.Cells(2, 7).Value = "東京都千代田区1-1-1"
objSheet.Cells(2, 8).Value = "'03-1234-5678"
objSheet.Cells(2, 9).Value = "Yes"
objSheet.Cells(2, 10).Value = "Yes"
objSheet.Cells(2, 11).Value = "MyCareLink Smart"
objSheet.Cells(2, 12).Value = Date()

Set objSheet = objWorkbook.Sheets("T_DeviceConfig")
objSheet.Cells(2, 1).Value = "C0001"
objSheet.Cells(2, 2).Value = "P0001"
objSheet.Cells(2, 3).Value = "2020/05/15"
objSheet.Cells(2, 4).Value = "Medtronic"
objSheet.Cells(2, 5).Value = "Azure XT DR MRI"
objSheet.Cells(2, 6).Value = "ABC123456"
objSheet.Cells(2, 7).Value = "RA-RV"
objSheet.Cells(2, 8).Value = "Active"

Set objSheet = objWorkbook.Sheets("T_Settings")
objSheet.Cells(2, 1).Value = "S0001"
objSheet.Cells(2, 2).Value = "P0001"
objSheet.Cells(2, 3).Value = "2020/05/15"
objSheet.Cells(2, 4).Value = "DDDR"
objSheet.Cells(2, 5).Value = 60
objSheet.Cells(2, 6).Value = 130
objSheet.Cells(2, 11).Value = 180

Set objSheet = objWorkbook.Sheets("T_Measurements")
objSheet.Cells(2, 1).Value = "M0001"
objSheet.Cells(2, 2).Value = "P0001"
objSheet.Cells(2, 3).Value = "2024/01/10"
objSheet.Cells(2, 4).Value = "外来"
objSheet.Cells(2, 5).Value = "3.0V"
objSheet.Cells(2, 6).Value = 450
objSheet.Cells(2, 7).Value = 520

objWorkbook.SaveAs strDbPath, 51
objWorkbook.Close False
objExcel.Quit

WScript.Echo "SUCCESS"

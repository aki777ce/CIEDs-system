Option Explicit
Dim fso, path, conn, rs, schema, sheetName
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName) & "\CIED_DB.xlsx"

Set conn = CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
    WScript.Quit
End If

Dim sheets
sheets = Array("T_Patients$", "T_DeviceConfig$", "T_Settings$", "T_Measurements$")

For Each sheetName In sheets
    WScript.Echo "--- SHEET: " & sheetName & " ---"
    Set rs = conn.Execute("SELECT TOP 1 * FROM [" & sheetName & "]")
    If Err.Number = 0 Then
        Dim i
        For i = 0 To rs.Fields.Count - 1
            WScript.Echo rs.Fields(i).Name
        Next
    Else
        WScript.Echo "Error reading sheet: " & Err.Description
    End If
    WScript.Echo ""
Next

conn.Close

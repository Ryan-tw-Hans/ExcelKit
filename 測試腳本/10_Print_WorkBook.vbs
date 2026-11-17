Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, workbookPath, librePath, printerName
status = "PASS"
message = "Print command success"
workbookPath = GetWorkbookPath()
librePath = GetLibreOfficePath()
printerName ="Test.xlsx" 'GetArgOrDefault(2, "")
msgbox printerName
On Error Resume Next
Set op = CreateOperator()
SaveMinimalWorkbook op

If Len(librePath) > 0 Then
    op.SetLibreOfficePath librePath
End If

op.PrintWorkbook workbookPath, printerName
msgbox printerName
If Err.Number <> 0 Then
    If InStr(1, LCase(Err.Description), "libreoffice", vbTextCompare) > 0 Then
        status = "SKIP"
        message = "LibreOffice not configured, skipped: " & Err.Description
    Else
        status = "FAIL"
        message = "Print failed: " & Err.Description
    End If
End If
On Error GoTo 0

LogTestResult "10", "Print Workbook", status, message
If status = "FAIL" Then WScript.Quit 1


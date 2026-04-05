Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, workbookPath, pdfPath, librePath
status = "PASS"
message = "Export PDF success"
workbookPath = GetWorkbookPath()
pdfPath = GetPdfPath()
librePath = GetLibreOfficePath()

On Error Resume Next
Set op = CreateOperator()
ResetTestFiles
SaveMinimalWorkbook op

If Len(librePath) > 0 Then
    op.SetLibreOfficePath librePath
End If

op.ExportToPdf workbookPath, pdfPath
If Err.Number <> 0 Then
    If InStr(1, LCase(Err.Description), "libreoffice", vbTextCompare) > 0 Then
        status = "SKIP"
        message = "LibreOffice not configured, skipped: " & Err.Description
    Else
        status = "FAIL"
        message = "Export PDF failed: " & Err.Description
    End If
ElseIf Not fso.FileExists(pdfPath) Then
    status = "FAIL"
    message = "PDF file not found: " & pdfPath
End If
On Error GoTo 0

LogTestResult "09", "Export To PDF", status, message
If status = "FAIL" Then WScript.Quit 1


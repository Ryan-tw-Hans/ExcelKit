Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, w
status = "PASS"
message = "OK"

On Error Resume Next
Set op = CreateOperator()
op.CreateNewWorkbook
op.SetColumnWidth "Sheet1", 2, 22.5
w = op.GetColumnWidth("Sheet1", 2)

If Err.Number <> 0 Then
    status = "FAIL"
    message = "ColumnWidth failed: " & Err.Description
ElseIf CDbl(w) < 22 Or CDbl(w) > 23 Then
    status = "FAIL"
    message = "Width verification failed: " & CStr(w)
End If
On Error GoTo 0

LogTestResult "18", "", status, message
If status = "FAIL" Then WScript.Quit 1


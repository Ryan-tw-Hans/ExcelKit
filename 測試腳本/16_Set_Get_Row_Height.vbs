Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, h
status = "PASS"
message = "OK"

On Error Resume Next
Set op = CreateOperator()
op.CreateNewWorkbook
op.SetRowHeight "Sheet1", 3, 25.5
h = op.GetRowHeight("Sheet1", 3)

If Err.Number <> 0 Then
    status = "FAIL"
    message = "RowHeight failed: " & Err.Description
ElseIf CDbl(h) < 25 Or CDbl(h) > 26 Then
    status = "FAIL"
    message = "Height verification failed: " & CStr(h)
End If
On Error GoTo 0

LogTestResult "16", "", status, message
If status = "FAIL" Then WScript.Quit 1


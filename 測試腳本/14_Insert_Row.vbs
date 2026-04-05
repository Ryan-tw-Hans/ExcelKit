Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, value
status = "PASS"
message = "OK"

On Error Resume Next
Set op = CreateOperator()
op.CreateNewWorkbook
op.SetCellValueByRC "Sheet1", 1, 1, "Top"
op.InsertRow "Sheet1", 1
value = op.GetCellValueByRC("Sheet1", 2, 1)

If Err.Number <> 0 Then
    status = "FAIL"
    message = "InsertRow failed: " & Err.Description
ElseIf CStr(value) <> "Top" Then
    status = "FAIL"
    message = "Shift verification failed, R2C1=" & CStr(value)
End If
On Error GoTo 0

LogTestResult "14", "", status, message
If status = "FAIL" Then WScript.Quit 1


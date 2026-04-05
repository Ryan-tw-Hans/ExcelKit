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
op.SetCellValueByRC "Sheet1", 1, 1, "ColA"
op.InsertColumn "Sheet1", 1
value = op.GetCellValueByRC("Sheet1", 1, 2)
If CStr(value) <> "ColA" Then
    status = "FAIL"
    message = "InsertColumn verification failed, R1C2=" & CStr(value)
Else
    op.DeleteColumn "Sheet1", 1
    value = op.GetCellValueByRC("Sheet1", 1, 1)
    If CStr(value) <> "ColA" Then
        status = "FAIL"
        message = "DeleteColumn verification failed, R1C1=" & CStr(value)
    End If
End If

If Err.Number <> 0 Then
    status = "FAIL"
    message = "Column operation failed: " & Err.Description
End If
On Error GoTo 0

LogTestResult "17", "", status, message
If status = "FAIL" Then WScript.Quit 1


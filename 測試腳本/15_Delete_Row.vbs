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
op.SetCellValueByRC "Sheet1", 1, 1, "Keep"
op.SetCellValueByRC "Sheet1", 2, 1, "DeleteMe"
op.DeleteRow "Sheet1", 2
value = op.GetCellValueByRC("Sheet1", 2, 1)

If Err.Number <> 0 Then
    status = "FAIL"
    message = "DeleteRow failed: " & Err.Description
ElseIf Not IsNull(value) And CStr(value) <> "" Then
    status = "FAIL"
    message = "Row 2 still has value: " & CStr(value)
End If
On Error GoTo 0

LogTestResult "15", "", status, message
If status = "FAIL" Then WScript.Quit 1


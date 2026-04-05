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
op.SetCellValue "Sheet1", "C3", "NeedClear"
op.ClearCell "Sheet1", "C3"
value = op.GetCellValue("Sheet1", "C3")

If Err.Number <> 0 Then
    status = "FAIL"
    message = "ClearCell failed: " & Err.Description
ElseIf Not IsNull(value) And CStr(value) <> "" Then
    status = "FAIL"
    message = "Cell still has value: " & CStr(value)
End If
On Error GoTo 0

LogTestResult "13", "", status, message
If status = "FAIL" Then WScript.Quit 1


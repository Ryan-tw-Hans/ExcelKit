Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op
status = "PASS"
message = "OK"

On Error Resume Next
Set op = CreateOperator()
op.CreateNewWorkbook
op.AddSheet "ToDelete"
op.DeleteSheet "ToDelete"

If Err.Number <> 0 Then
    status = "FAIL"
    message = "DeleteSheet failed: " & Err.Description
End If
On Error GoTo 0

LogTestResult "12", "", status, message
If status = "FAIL" Then WScript.Quit 1


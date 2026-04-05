Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op
status = "PASS"
message = "?êÂ?Âª∫Á? COM ?©‰ª∂ ExcelKit.ExcelOperator"

On Error Resume Next
Set op = CreateOperator()
If Err.Number <> 0 Or (op Is Nothing) Then
    status = "FAIL"
    message = "?°Ê?Âª∫Á? COM ?©‰ª∂: " & Err.Description
End If
On Error GoTo 0

LogTestResult "01", "ÂºïÁî®COM?É‰ª∂", status, message
If status = "FAIL" Then WScript.Quit 1


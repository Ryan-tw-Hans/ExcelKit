Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, dllVersion, comVersion
status = "PASS"

On Error Resume Next
Set op = CreateOperator()
comVersion = op.GetVersionInfo()
If Err.Number <> 0 Then
    status = "FAIL"
    message = "è®€??COM ?ˆæœ¬å¤±æ?: " & Err.Description
Else
    dllVersion = GetDllVersion()
    message = "DLL Version=" & dllVersion & "; COM Version=" & comVersion
End If
On Error GoTo 0

LogTestResult "02", "è®€?–DLL?ˆæœ¬", status, message
If status = "FAIL" Then WScript.Quit 1


Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, librePath, logDir
status = "PASS"
message = "Config functions success"
librePath = GetLibreOfficePath()
logDir = BuildPath(GetDataFolder(), "Logs")
EnsureFolder logDir

On Error Resume Next
Set op = CreateOperator()

If Len(librePath) > 0 Then
    op.SetLibreOfficePath librePath
End If
op.SetLogDirectory logDir
op.SetLogEnabled True
op.SetLogLevel 3

If Err.Number <> 0 Then
    status = "FAIL"
    message = "Config function failed: " & Err.Description
End If
On Error GoTo 0

LogTestResult "20", "Config Functions", status, message
If status = "FAIL" Then WScript.Quit 1


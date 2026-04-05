Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, workbookPath
status = "PASS"
message = "Create + save workbook success"
workbookPath = GetWorkbookPath()

On Error Resume Next
Set op = CreateOperator()
ResetTestFiles
op.CreateNewWorkbook
op.SaveWorkbook workbookPath
If Err.Number <> 0 Then
    status = "FAIL"
    message = "Create/Save workbook failed: " & Err.Description
ElseIf Not fso.FileExists(workbookPath) Then
    status = "FAIL"
    message = "Workbook file not found after save: " & workbookPath
End If
On Error GoTo 0

LogTestResult "03", "Create Workbook", status, message
If status = "FAIL" Then WScript.Quit 1


Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, value, workbookPath, expectedValue
status = "PASS"
message = "Set/Get by cell name success"
workbookPath = GetWorkbookPath()
expectedValue = GetDynamicValue("HelloByName")

On Error Resume Next
Set op = CreateOperator()
ResetTestFiles
op.CreateNewWorkbook
op.SetCellValue "Sheet1", "A1", expectedValue
op.SaveWorkbook workbookPath

value = ReadCellValueByAdo(workbookPath, "Sheet1", "A1")
msgbox value
If Err.Number <> 0 Then
    status = "FAIL"
    message = "Set/Get by name failed: " & Err.Description
ElseIf CStr(value) <> expectedValue Then
    status = "FAIL"
    message = "Verification failed, A1=" & CStr(value) & ", expected=" & expectedValue
End If
On Error GoTo 0

LogTestResult "05", "Set Cell By Name", status, message
If status = "FAIL" Then WScript.Quit 1


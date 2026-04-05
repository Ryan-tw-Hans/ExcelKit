Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, value, workbookPath, expectedValue
status = "PASS"
message = "Set/Get by row column success"
workbookPath = GetWorkbookPath()
expectedValue = GetDynamicValue("HelloByRC")

On Error Resume Next
Set op = CreateOperator()
ResetTestFiles
op.CreateNewWorkbook
op.SetCellValueByRC "Sheet1", 2, 3, expectedValue
op.SaveWorkbook workbookPath

value = ReadCellValueByAdo(workbookPath, "Sheet1", GetCellAddressByRC(2, 3))

If Err.Number <> 0 Then
    status = "FAIL"
    message = "Set/Get by row column failed: " & Err.Description
ElseIf CStr(value) <> expectedValue Then
    status = "FAIL"
    message = "Verification failed, R2C3=" & CStr(value) & ", expected=" & expectedValue
End If
On Error GoTo 0

LogTestResult "06", "Set Cell By RC", status, message
If status = "FAIL" Then WScript.Quit 1


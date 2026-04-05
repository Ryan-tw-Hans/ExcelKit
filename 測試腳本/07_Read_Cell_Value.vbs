Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, verifyOp, value, workbookPath, expectedValue
status = "PASS"
message = "Read cell value success"
workbookPath = GetWorkbookPath()
expectedValue = GetDynamicValue("ReadBack")

On Error Resume Next
Set op = CreateOperator()
ResetTestFiles
op.CreateNewWorkbook
op.SetCellValue "Sheet1", "B2", expectedValue
op.SaveWorkbook workbookPath

Set verifyOp = CreateOperator()
verifyOp.OpenWorkbook workbookPath
value = verifyOp.GetCellValue("Sheet1", "B2")

If Err.Number <> 0 Then
    status = "FAIL"
    message = "Read cell value failed: " & Err.Description
ElseIf CStr(value) <> expectedValue Then
    status = "FAIL"
    message = "Verification failed, B2=" & CStr(value) & ", expected=" & expectedValue
End If
On Error GoTo 0

LogTestResult "07", "Read Cell Value", status, message
If status = "FAIL" Then WScript.Quit 1


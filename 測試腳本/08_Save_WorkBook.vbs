Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op, verifyOp, workbookPath, expectedValue, value
status = "PASS"
message = "Save workbook success"
workbookPath = GetWorkbookPath()
expectedValue = GetDynamicValue("SaveTest")

On Error Resume Next
Set op = CreateOperator()
ResetTestFiles
op.CreateNewWorkbook
op.SetCellValue "Sheet1", "A1", expectedValue
op.SaveWorkbook workbookPath

Set verifyOp = CreateOperator()
verifyOp.OpenWorkbook workbookPath
value = verifyOp.GetCellValue("Sheet1", "A1")

If Err.Number <> 0 Then
    status = "FAIL"
    message = "Save workbook failed: " & Err.Description
ElseIf Not fso.FileExists(workbookPath) Then
    status = "FAIL"
    message = "Workbook file not found after save: " & workbookPath
ElseIf CStr(value) <> expectedValue Then
    status = "FAIL"
    message = "Saved value mismatch, A1=" & CStr(value) & ", expected=" & expectedValue
End If
On Error GoTo 0

LogTestResult "08", "Save Workbook", status, message
If status = "FAIL" Then WScript.Quit 1


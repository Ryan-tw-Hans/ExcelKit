Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "_TestCommon.vbs"), 1).ReadAll

Dim status, message, op1, op2, workbookPath, value, expectedValue
status = "PASS"
message = "Open workbook success"
workbookPath = GetWorkbookPath()
expectedValue = GetDynamicValue("OpenTest")

On Error Resume Next
Set op1 = CreateOperator()
ResetTestFiles
op1.CreateNewWorkbook
op1.SetCellValue "Sheet1", "A1", expectedValue
op1.SaveWorkbook workbookPath

Set op2 = CreateOperator()
op2.OpenWorkbook workbookPath
value = op2.GetCellValue("Sheet1", "A1")

If Err.Number <> 0 Then
    status = "FAIL"
    message = "Open workbook failed: " & Err.Description
ElseIf CStr(value) <> expectedValue Then
    status = "FAIL"
    message = "Content verification failed, A1=" & CStr(value) & ", expected=" & expectedValue
End If
On Error GoTo 0

LogTestResult "11", "Open Workbook", status, message
If status = "FAIL" Then WScript.Quit 1


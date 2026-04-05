' ExcelKit end-to-end test (hardened)
On Error Resume Next

Dim excel, fso, shell, baseDir, filePath, pdfPath, logDir, libreOfficePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Use TEMP folder to avoid failures when D:\ does not exist
baseDir = shell.ExpandEnvironmentStrings("%TEMP%") & "\ExcelKitTest"
If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir

filePath = baseDir & "\Test.xlsx"
pdfPath = baseDir & "\TestExcel.pdf"
logDir = baseDir & "\Logs"
If Not fso.FolderExists(logDir) Then fso.CreateFolder logDir

libreOfficePath = "D:\00-program file\LibreOfficePortable\App\libreoffice\program\soffice.exe"

' 1. Create COM object
Set excel = CreateObject("ExcelKit.ExcelOperator")
CheckErr "Create ExcelKit.ExcelOperator COM object"

' 2. Optional LibreOffice path override
' If this path is missing, skip and rely on auto-detection
If fso.FileExists(libreOfficePath) Then
    excel.SetLibreOfficePath libreOfficePath
    CheckErr "Set LibreOffice path"
Else
    Notify "Info: configured LibreOffice path not found, using auto-detection."
End If

' 3. Configure logging
excel.SetLogEnabled True
CheckErr "Enable logging"
excel.SetLogDirectory logDir
CheckErr "Set log directory"
excel.SetLogLevel 3
CheckErr "Set log level"

' 4. Create workbook
excel.CreateNewWorkbook
CheckErr "Create workbook"

' 5. Write sample data
excel.SetCellValue "Sheet2", "A1", "Hello ExcelKit!"
CheckErr "Write A1"
excel.SetCellValue "Sheet2", "A2", 123.45
CheckErr "Write A2"
excel.SetCellValue "Sheet2", "B1", "Date: " & Now()
CheckErr "Write B1"
excel.SetCellValue "Sheet2", "B2", "01234"
CheckErr "Write B2"
excel.SetCellValueByRC "Sheet2", 5, 5, 123.45
CheckErr "Write R5C5"

' 6. Get version
Dim version
version = excel.GetVersionInfo()
CheckErr "Get version info"
Notify "Component version: " & version

' 7. Save workbook
excel.SaveWorkbook filePath
CheckErr "Save workbook"
Notify "Excel file saved: " & filePath

' 8. Export to PDF (uses LibreOffice)
excel.ExportToPdf filePath, pdfPath
CheckErr "Export to PDF"
Notify "PDF exported: " & pdfPath

' 9. Print workbook (requires LibreOffice)
' excel.PrintWorkbook filePath, "" ' ‰ΩøÁî®?êË®≠?∞Ë°®Ê©?

' 10. Release resources
Set excel = Nothing
Notify "Test completed."
WScript.Quit 0

Sub CheckErr(stepName)
    If Err.Number <> 0 Then
        Dim detail
        detail = "Step failed: " & stepName & vbCrLf & _
                 "Err.Number: " & Err.Number & " (0x" & Hex(Err.Number) & ")" & vbCrLf & _
             "Description: " & Err.Description
        Notify detail
        WScript.Quit 1
    End If
    Err.Clear
End Sub

Sub Notify(message)
    If InStr(1, LCase(WScript.FullName), "cscript.exe", vbTextCompare) > 0 Then
        WScript.Echo message
    Else
        MsgBox message
    End If
End Sub


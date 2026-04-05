' ExcelKit end-to-end test (hardened)
On Error Resume Next

Dim excel, fso, shell, baseDir, filePath, pdfPath, logDir, libreOfficePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")



filePath = "D:\Test.xlsx"
pdfPath = "D:\TestExcel.pdf"
logDir = "D:\Logs"
If Not fso.FolderExists(logDir) Then fso.CreateFolder logDir

libreOfficePath = "D:\00-program file\LibreOfficePortable\App\libreoffice\program\soffice.exe"

' 1. Create COM object
Set excel = CreateObject("ExcelKit.ExcelOperator")
CheckErr "Create ExcelKit.ExcelOperator COM object"
Notify "ExcelKit version: " & excel.GetVersionInfo()
CheckErr "Get version info"

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

' 4. ËΩâÊ??¢Ê?Ê™îÊ?ÔºöÂ? OpenWorkbookÔºåÂ? ExportToPdf
If Not fso.FileExists(filePath) Then
    Notify "Step failed: Source workbook not found: " & filePath
    WScript.Quit 1
End If

excel.OpenWorkbook filePath
CheckErr "Open workbook"



' 8. Export to PDF (uses LibreOffice)
excel.ExportToPdf filePath, pdfPath
CheckErr "Export to PDF"
Notify "PDF exported: " & pdfPath

' 9. Print workbook (requires LibreOffice)
 excel.PrintWorkbook filePath, "" ' ‰ΩøÁî®?êË®≠?∞Ë°®Ê©?

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


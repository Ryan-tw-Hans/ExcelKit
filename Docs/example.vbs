' ExcelKit 總體測試範例
On Error Resume Next

Dim excel, filePath
filePath = "C:\Temp\TestExcel.xlsx" ' 請確保該目錄存在

' 1. 建立物件
Set excel = CreateObject("ExcelKit.ExcelOperator")

If Err.Number <> 0 Then
    MsgBox "無法建立 ExcelKit 物件，請檢查是否已成功註冊。" & vbCrLf & Err.Descript
    WScript.Quit
End If

' 2. (選填) LibreOffice 路徑設定
' 元件會自動嘗試尋找 LibreOffice，但若您的安裝路徑較為特殊，建議手動設定路徑
excel.SetLibreOfficePath "C:\Users\ryan_\source\repos\Excelopen\測試腳本\LibreOfficePortable\App\libreoffice\program\soffice.exe"

' 3. 設定日誌
excel.SetLogEnabled True
excel.SetLogDirectory "C:\Logs\ExcelKit"
excel.SetLogLevel 3 ' 詳細記錄
' 4. 建立新的活頁簿
excel.CreateNewWorkbook
' 5. 寫入資料到指定的 Sheet 和儲存格
excel.SetCellValue "Sheet1", "A1", "Hello ExcelKit!"
excel.SetCellValue "Sheet1", "A2", 123.45
excel.SetCellValue "Sheet1", "B1", "日期：" & Now()
' 6. 取得版本資訊
Dim version
version = excel.GetVersionInfo()
MsgBox "目前元件版本: " & version
' 7. 儲存檔案
excel.SaveWorkbook filePath
if Err.Number <> 0 Then
    MsgBox "發生錯誤: " & Err.Description
Else
    MsgBox "Excel 檔案已成功儲存至: " & filePath
End If
' 8. 匯出為 PDF (會自動調用 LibreOffice)
Dim pdfPath
pdfPath = "C:\Temp\TestExcel.pdf"
excel.ExportToPdf filePath, pdfPath
If Err.Number <> 0 Then
    MsgBox "匯出 PDF 發生錯誤: " & Err.Description
Else
    MsgBox "Excel 檔案已成功匯出為 PDF: " & pdfPath
End If
' 9. 列印活頁簿 (這需要 LibreOffice)
' excel.PrintWorkbook filePath, "" ' 使用預設印表機
' 10. 釋放資源
Set excel = Nothing
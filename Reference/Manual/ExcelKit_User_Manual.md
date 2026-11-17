# ExcelKit COM 元件使用手冊

本手冊提供 ExcelKit (原 Excelopen) COM 元件的註冊、卸載、驗證及 VBScript 使用範例說明。

## 目錄
1. [手動註冊元件 (Manual Registration)](#1-手動註冊元件-manual-registration)
2. [卸載元件 (Unregistration)](#2-卸載元件-unregistration)
3. [確認註冊狀態](#3-確認註冊狀態)
4. [API 快速檢索表 (Quick API Index)](#4-api-快速檢索表-quick-api-index)
5. [API 函式與屬性詳細說明](#5-api-函式與屬性詳細說明)
    - [5.1 基礎操作](#51-基礎操作-basic-operations)
    - [5.2 資料編輯](#52-資料編輯-data-manipulation)
    - [5.3 行與列的操作](#53-行與列的操作-rows-and-columns)
    - [5.4 擴充功能](#54-擴充功能-extended-features---需要-libreoffice)
    - [5.5 配置與日誌](#55-配置與日誌-configuration--logging)
6. [常用數值對照 (Constants Reference)](#6-常用數值對照-constants-reference)
7. [VBScript (VBS) 綜合測試腳本](#7-vbscript-vbs-綜合測試腳本)
8. [常見問題 (FAQ)](#8-常見問題-faq)

## 1. 手動註冊元件 (Manual Registration)

註冊 COM 元件需要使用管理員權限打開「命令提示字元 (CMD)」或「PowerShell」。

### ⚠️ 注意事項
* 請確保已安裝 **.NET Framework 4.7.2**。
* 執行路徑必須指向編譯後的 `ExcelKit.dll` 實際位置。
* 必須以 **管理員身份** 執行。

### 64 位元 (x64) 註冊
適用於 64 位元應用程式 (如 64 位元 LabVIEW, Excel, VBScript x64)。
```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase "你的路徑\ExcelKit.dll" /tlb "你的路徑\ExcelKit.tlb"
```

### 32 位元 (x86) 註冊
適用於 32 位元應用程式 (如 32 位元 LabVIEW, Excel, VBScript x32)。
```cmd
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe /codebase "你的路徑\ExcelKit.dll" /tlb "你的路徑\ExcelKit.tlb"
```

---

## 2. 卸載元件 (Unregistration)

如果需要更新元件或移除，可執行以下指令：

### 64 位元 (x64) 卸載
```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /unregister "你的路徑\ExcelKit.dll"
```

### 32 位元 (x86) 卸載
```cmd
C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe /unregister "你的路徑\ExcelKit.dll"
```

---

## 3. 確認註冊狀態

### 方法 A：使用 PowerShell 測試 (推薦)
開啟 PowerShell 並執行以下指令，若無報錯並顯示版本資訊，表示註冊成功：
```powershell
$excel = New-Object -ComObject ExcelKit.ExcelOperator
$excel.GetVersionInfo()
```
或者直接執行產送的 `check-registration.ps1` 檔案：
```powershell
.\check-registration.ps1
```

### 方法 B：檢查登錄檔 (Registry Editor)
1. 按 `Win + R`，輸入 `regedit` 並執行。
2. 尋找路徑：`HKEY_CLASSES_ROOT\ExcelKit.ExcelOperator`。
3. 如果該路徑存在，表示 ProgID 已載入。

---

## 4. API 快速檢索表 (Quick API Index)

| 類型 | 名稱 | 簡要說明 | 詳情章節 |
| :--- | :--- | :--- | :--- |
| **函式** | `GetVersionInfo()` | 取得元件版本資訊字串 | [5.1](#getversioninfo) |
| **函式** | `CreateNewWorkbook()` | 在記憶體中建立新活頁簿 | [5.1](#createnewworkbook) |
| **函式** | `OpenWorkbook(path)` | 開啟現有 Excel 檔案 | [5.1](#openworkbookfilepath) |
| **函式** | `SaveWorkbook(path)` | 儲存目前活頁簿 (可另存) | [5.1](#saveworkbookfilepath) |
| **函式** | `AddSheet(name)` | 新增一個工作表 | [5.2](#addsheetsheetname) |
| **函式** | `DeleteSheet(name)` | 刪除指定名稱的工作表 | [5.2](#deletesheetsheetname) |
| **函式** | `SetCellValue(sh, addr, val)` | 使用 A1 格式寫入資料 | [5.2](#setcellvaluesheetname-celladdress-value) |
| **函式** | `SetCellValueByRC(sh, r, c, val)` | 使用行列號 (1-based) 寫入資料 | [5.2](#setcellvaluebyrcsheetname-row-column-value) |
| **函式** | `GetCellValue(sh, addr)` | 使用 A1 格式讀取資料 | [5.2](#getcellvaluesheetname-celladdress) |
| **函式** | `ClearCell(sh, addr)` | 清除單一儲存格內容 | [5.2](#clearcellsheetname-celladdress) |
| **函式** | `InsertRow/Column(sh, pos)` | 插入空行/空列 | [5.3](#53-行與列的操作-rows-and-columns) |
| **函式** | `SetRow/ColumnWidth(sh, p, v)` | 設定指定行列寬高 | [5.3](#53-行與列的操作-rows-and-columns) |
| **函式** | `ExportToPdf(src, dst)` | 將 Excel 轉換為 PDF (需 LibreOffice) | [5.4](#54-擴充功能-extended-features---需要-libreoffice) |
| **函式** | `PrintWorkbook(src, prn)` | 列印 Excel 檔案 (需 LibreOffice) | [5.4](#54-擴充功能-extended-features---需要-libreoffice) |
| **函式** | `SetLibreOfficePath(path)` | 手動指定 LibreOffice 執行路徑 | [5.5](#55-配置與日誌-configuration--logging) |
| **函式** | `SetLogEnabled(bool)` | 啟用或停用操作日誌 | [5.5](#55-配置與日誌-configuration--logging) |
| **常數** | `LogLevelError (1)` | 僅記錄錯誤 | [6.0](#6-常用數值對照-constants-reference) |
| **常數** | `LogLevelInfo (3)` | 詳細執行記錄 (預設) | [6.0](#6-常用數值對照-constants-reference) |

---

## 5. API 函式與屬性詳細說明

以下是 `ExcelKit.ExcelOperator` 物件支援的所有公開介面說明。

### 5.1 基礎操作 (Basic Operations)

#### `GetVersionInfo()`
*   **說明**: 取得目前的元件名稱、版本與更新日期。
*   **回傳**: String
*   **VBS 範例**:
    ```vbscript
    MsgBox excel.GetVersionInfo()
    ```

#### `CreateNewWorkbook()`
*   **說明**: 在記憶體中建立一個新的活頁簿（預設包含一個 "Sheet1" 分頁）。
*   **VBS 範例**:
    ```vbscript
    excel.CreateNewWorkbook
    ```

#### `OpenWorkbook(filePath)`
*   **說明**: 開啟現有的 Excel 檔案。
*   **參數**: `filePath` (String) - 檔案完整路徑。
*   **VBS 範例**:
    ```vbscript
    excel.OpenWorkbook "D:\Reports\Data.xlsx"
    ```

#### `SaveWorkbook(filePath)`
*   **說明**: 儲存目前的活頁簿。
*   **參數**: `filePath` (String, 可省略) - 儲存位置。若省略則覆蓋原開啟檔案。
*   **VBS 範例**:
    ```vbscript
    ' 另存新檔
    excel.SaveWorkbook "C:\Output\Result.xlsx"
    ' 儲存原檔
    excel.SaveWorkbook ""
    ```

---

### 5.2 資料編輯 (Data Manipulation)

#### `AddSheet(sheetName)`
*   **說明**: 新增一個指定名稱的工作表。
*   **VBS 範例**:
    ```vbscript
    excel.AddSheet "NewReport"
    ```

#### `SetCellValue(sheetName, cellAddress, value)`
*   **說明**: 在指定分頁的儲存格寫入資料。支援 A1 格式 (如 "B2")。
*   **VBS 範例**:
    ```vbscript
    excel.SetCellValue "Sheet1", "A1", "測試資料"
    ```

#### `SetCellValueByRC(sheetName, row, column, value)`
*   **說明**: 使用行列號寫入資料。行列號均為 1-based (從 1 開始)。
*   **VBS 範例**:
    ```vbscript
    ' 在第 2 行第 3 列 (C2) 寫入內容
    excel.SetCellValueByRC "Sheet1", 2, 3, "RC模式"
    ```

#### `GetCellValue(sheetName, cellAddress)`
*   **說明**: 取得指定儲存格的內容。
*   **回傳**: Variant (根據內容返回對應類型)。

#### `ClearCell(sheetName, cellAddress)`
*   **說明**: 清除指定區域或儲存格的內容，保留格式。
*   **VBS 範例**:
    ```vbscript
    excel.ClearCell "Sheet1", "A1"
    ```

#### `DeleteSheet(sheetName)`
*   **說明**: 刪除指定的工作表。

#### `GetAllSheetNames()`
*   **說明**: 取得活頁簿中所有工作表的名稱列表。
*   **回傳**: Array (字串陣列)
*   **VBS 範例**:
    ```vbscript
    names = excel.GetAllSheetNames()
    For Each name In names
        WScript.Echo "工作表: " & name
    Next
    ```

---

### 5.3 行與列的操作 (Rows and Columns)

#### `InsertRow(sheetName, rowNumber)` / `DeleteRow(sheetName, rowNumber)`
*   **說明**: 在指定位置插入空列或刪除該行。
*   **VBS 範例**:
    ```vbscript
    excel.InsertRow "Sheet1", 5 ' 在第 5 行上方插入一空行
    excel.DeleteRow "Sheet1", 10 ' 刪除第 10 行
    ```

#### `SetRowHeight(sh, row, height)` / `SetColumnWidth(sh, col, width)`
*   **說明**: 設定指定行的高或列的寬。
*   **VBS 範例**:
    ```vbscript
    excel.SetRowHeight "Sheet1", 1, 30.5
    excel.SetColumnWidth "Sheet1", 1, 20.0
    ```

---

### 5.4 擴充功能 (Extended Features - 需要 LibreOffice)

這組功能需要系統安裝 [LibreOffice](https://zh-tw.libreoffice.org/)。元件會自動探測路徑，若失敗可使用 `SetLibreOfficePath` 手動指定。

LibreOffice 官方下載連結
*   https://zh-tw.libreoffice.org/download/download-libreoffice/

LibreOffice 官方免安裝版下載連結
*   https://zh-tw.libreoffice.org/download/portable-versions/

#### `ExportToPdf(excelFilePath, pdfOutputPath)`
*   **說明**: 將指定的 Excel 檔案轉換為 PDF 格式。
*   **參數**:
    - `excelFilePath`: 來源 Excel 路徑。
    - `pdfOutputPath`: PDF 輸出路徑（包含檔名）。
*   **VBS 範例**:
    ```vbscript
    excel.ExportToPdf "C:\Data\Report.xlsx", "C:\Data\Report.pdf"
    ```

#### `PrintWorkbook(excelFilePath, printerName)`
*   **說明**: 使用 LibreOffice 列印 Excel 檔案。
*   **參數**:
    - `excelFilePath`: 欲列印的檔案。
    - `printerName`: 印表機名稱。若設為空字串 `""` 則使用系統預設印表機。
*   **VBS 範例**:
    ```vbscript
    ' 使用預設印表機
    excel.PrintWorkbook "C:\Data\Report.xlsx", ""
    ' 指定印表機
    excel.PrintWorkbook "C:\Data\Report.xlsx", "HP LaserJet 400"
    ```

---

### 5.5 配置與日誌 (Configuration & Logging)

#### `SetLibreOfficePath(libreOfficePath)`
*   **說明**: 手動設定 `soffice.exe` 的路徑（例如當安裝在非標準路徑時）。

#### `SetLogEnabled(enabled)` / `SetLogDirectory(path)` / `SetLogLevel(level)`
*   **說明**: 控制操作日誌系統。
*   **參數**: Level 1=Error, 2=Warning, 3=Info。
*   **VBS 範例**:
    ```vbscript
    excel.SetLogEnabled True
    excel.SetLogDirectory "C:\Logs\ExcelKit"
    excel.SetLogLevel 3 ' 詳細記錄
    ```

---

## 6. 常用數值對照 (Constants Reference)

雖然 VBScript 無法直接存取 DLL 內部的常數，但在呼叫 `SetLogLevel` 等函式時，請參考以下對照表：

| 功能項目 | 建議數值 | 說明 |
| :--- | :--- | :--- |
| **LogLevel - Error** | `1` | 僅記錄發生中斷的錯誤。 |
| **LogLevel - Warning**| `2` | 記錄錯誤與非預期的警告訊息。 |
| **LogLevel - Info** | `3` | (預設) 詳細記錄每一步操作與執行狀態。 |

---

## 7. VBScript (VBS) 綜合測試腳本

以下是一個完整的 VBScript 測試範例，涵蓋了 ExcelKit 的主要功能：

```vbscript
' ExcelKit comprehensive test example
On Error Resume Next

Dim excel, filePath
filePath = "C:\Temp\TestExcel.xlsx" ' Make sure this directory exists

' 1. Create object
Set excel = CreateObject("ExcelKit.ExcelOperator")

If Err.Number <> 0 Then
    MsgBox "Unable to create ExcelKit object. Please verify that registration completed successfully." & vbCrLf & Err.Description
    WScript.Quit
End If

' 2. (Optional) LibreOffice path setting
' The component will try to find LibreOffice automatically, but if your install path is non-standard, set it manually
' excel.SetLibreOfficePath "C:\Program Files\LibreOffice\program\soffice.exe"

' 3. Configure logging
excel.SetLogEnabled True
excel.SetLogDirectory "C:\Logs\ExcelKit"
excel.SetLogLevel 3 ' Verbose logging

' 4. Create a new workbook
excel.CreateNewWorkbook

' 5. Write data to specific sheet and cells
excel.SetCellValue "Sheet1", "A1", "Hello ExcelKit!"
excel.SetCellValue "Sheet1", "A2", 123.45
excel.SetCellValue "Sheet1", "B1", "Date: " & Now()

' 6. Get version information
Dim version
version = excel.GetVersionInfo()
MsgBox "Current component version: " & version

' 7. Save file
excel.SaveWorkbook filePath

if Err.Number <> 0 Then
    MsgBox "Error occurred: " & Err.Description
Else
    MsgBox "Excel file saved successfully to: " & filePath
End If

' 8. Export to PDF (LibreOffice will be invoked automatically)
Dim pdfPath
pdfPath = "C:\Temp\TestExcel.pdf"
excel.ExportToPdf filePath, pdfPath

If Err.Number <> 0 Then
    MsgBox "Error exporting PDF: " & Err.Description
Else
    MsgBox "Excel file exported to PDF successfully: " & pdfPath
End If

' 9. Print workbook (requires LibreOffice)
' excel.PrintWorkbook filePath, "" ' Use default printer

' 10. Release resources
Set excel = Nothing
```

---

## 8. 常見問題 (FAQ)

*   **錯誤碼 0x80040154 (Class not registered)**：
    *   通常是因為 DLL 沒有被註冊，或者你的程式位元數 (x86/x64) 與註冊的 DLL 位元數不對應。請確保同時執行過 x86 與 x64 的 `regasm` 指令。
*   **權限不足**：
    *   註冊時必須對 CMD 點選「右鍵 -> 以管理員身分執行」。
*   **FileNotFoundException**：
    *   使用 `regasm` 時，請務必加上 `/codebase` 參數，否則 .NET 會去 GAC 尋找 DLL。

## 儲存格文字顏色（RGB）

使用 `SetCellFontColor(sheetName, cellAddress, red, green, blue)` 依位址設定，或使用 `SetCellFontColorByRC(sheetName, row, column, red, green, blue)` 依行列設定；行、列從 1 開始。RGB 三個整數各為 0–255，超出範圍會回報錯誤；工作表必須已存在。只設定整個儲存格的文字顏色，不改變儲存格值或其他格式。儲存活頁簿後顏色才會寫入檔案。

```vbscript
excel.SetCellValue "Sheet1", "A1", "紅色文字"
excel.SetCellFontColor "Sheet1", "A1", 255, 0, 0
excel.SetCellValueByRC "Sheet1", 2, 2, "藍色文字"
excel.SetCellFontColorByRC "Sheet1", 2, 2, 0, 0, 255
excel.SaveWorkbook filePath
```

完整 PowerShell 範例請見 `Examples/PowerShell/FontColor.ps1`。不支援局部字元、背景色或主題色。

## DLL 版本與 COM 引用

COM 呼叫名稱維持 `ExcelKit.ExcelOperator`，例如 `CreateObject("ExcelKit.ExcelOperator")`。相容更新維持組件版本與現有介面識別，並以獨立新增介面方式擴充；仍需對實際用戶端驗證相容性。

DLL 的「檔案版本」使用 `正式基準.修訂序號`（例如 `1.2.4.1`），可在 Windows 檔案內容的詳細資料中查看；每次程式更新採用比前次更大的版本。產品版本例如 `1.2.4+20261116223400` 使用時間戳，COM 組件版本 `1.2.4.0` 是不同欄位，請勿以組件版本判斷檔案是否相同。

開發者發布前須更新元件產品版本、資訊版本、Class1 中的版本和 DLL FileVersion，再執行 `powershell.exe -NoProfile -File Installer/FileVersion.ps1 -Mode Reserve` 預留；接著打包預檢會比對台帳、來源、既有報告與可讀的本機 COM 登錄 DLL。版本衝突或資訊無法讀取時停止。版本台帳須在發行者間同步協調；未納入的其他機器或檔案不會自動被發現。已預留的版本不回收，來源變動要重新分配更大版本與新產品時間戳。

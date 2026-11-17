# 軟體需求規格書（SRS）｜ExcelKit v1.2.4

| 項目 | 內容 |
| --- | --- |
| 文件編號 | SRS-ExcelKit-v1.2.4 |
| 文件版本 | v0.1（草稿，待人類確認） |
| 撰寫日期 | 2026-11-14（Asia/Taipei） |
| 撰寫者 | Claude（PM／SA 角色，依現況程式碼反向整理；自查，非獨立審查） |
| 依據來源 | `Excelopen/Class1.cs`、`Excelopen/Excelopen.csproj`、`Excelopen/Properties/AssemblyInfo.cs`、`README.MD`、`Reference/Manual/ExcelKit_User_Manual.md`、`測試腳本/*.vbs` |
| 受測／描述版本 | 程式版本 1.2.4（csproj `Version`、`ExcelOperator.Version`），最後更新日期常數 2026-04-05 |
| 相關文件 | `TestPlans/STP-ExcelKit-v1.2.4.md`、`TestPlans/TCS-ExcelKit-v1.2.4.md`、`TestPlans/RTM-ExcelKit-v1.2.4.md`、`Requirements/CodeReview-ExcelKit-v1.2.4.md` |

> 本文件為「現況規格（as-is）」：描述目前程式碼實際行為。凡實際行為與手冊或合理預期不一致者，於各需求「備註」標示並以 `ISS-xx` 連到問題清單，**不在此直接改寫為期望行為**；是否修正由人類決策。

---

## 1. 簡介

### 1.1 目的
ExcelKit 是以 C# 開發的 Excel（.xlsx）操作元件，主要以 COM 方式提供 LabVIEW、VBScript 及其他 COM Client 呼叫，在不安裝 Microsoft Excel 的情況下建立、讀寫、儲存活頁簿，並透過外部 LibreOffice 進行 PDF 匯出與列印。

### 1.2 範圍
- 包含：`ExcelKit.ExcelOperator` COM 類別的 26 個 COM 方法、.NET 端額外公開方法、LibreOffice 設定與日誌機制、建置與部署限制。
- 不包含：LibreOffice 本身的轉換品質、印表機驅動、LabVIEW VI 內部邏輯、`agent-workflow/`／`AgentUse/` 等 Agent 控制資料。

### 1.3 名詞定義
| 名詞 | 說明 |
| --- | --- |
| COM Client | 透過 ProgID 建立物件的呼叫端（LabVIEW、VBScript、PowerShell 等） |
| 活頁簿（Workbook） | 記憶體中的 ClosedXML `XLWorkbook` 物件 |
| 工作表（Sheet） | 活頁簿中的 Worksheet |
| A1 位址 | 以欄字母＋列號表示的儲存格位址，例如 `B2` |
| RC | Row／Column 數字索引，1-based |
| soffice | LibreOffice 執行檔 `soffice.exe` |

### 1.4 系統概觀
```
COM Client (LabVIEW / VBS / PowerShell)
        │  IDispatch (ProgID: ExcelKit.ExcelOperator)
        ▼
ExcelOperator（ExcelKit.dll，net472，Costura 內嵌相依）
   ├─ ClosedXML 0.105.0  ── 讀寫 .xlsx（記憶體中）
   ├─ LibreOfficeConfigManager（static，行程內共用）
   │     ├─ LibreOffice 路徑解析
   │     └─ 日誌寫檔
   └─ Process.Start(soffice --headless ...) ── PDF 匯出／列印
```

---

## 2. 整體描述

### 2.1 產品屬性
| 項目 | 內容 |
| --- | --- |
| 組件名稱 | `ExcelKit`（AssemblyName／RootNamespace），C# 命名空間實際為 `Excelopen` |
| 目標框架 | `net472`（COM 使用、Costura 內嵌相依）、`net10.0`（.NET 函式庫用途，停用 Fody） |
| COM ProgID | `ExcelKit.ExcelOperator` |
| COM 類別 GUID | `87654321-4321-4321-4321-210987654321` |
| COM 介面 GUID | `12345678-1234-1234-1234-123456789012`（`IExcelOperator`，IDispatch） |
| 組件 GUID | `5F4B3D3E-6C2F-4F77-9D65-4B8A25E1B9A1` |
| 主要相依 | ClosedXML 0.105.0、Fody 6.8.1、Costura.Fody 5.7.0、System.Diagnostics.DiagnosticSource 8.0.0 |
| 外部相依 | LibreOffice（僅 PDF／列印需要） |
| 授權 | MIT |

### 2.2 使用者與使用情境
| 使用者 | 典型情境 |
| --- | --- |
| LabVIEW 量測程式 | 量測完成後把結果寫入 Excel 範本、存檔、轉 PDF 或直接列印（`Reference/LabVIEW/*.vi`） |
| VBScript／批次自動化 | 產生報表、讀取既有檔案資料 |
| .NET 開發者（net10.0） | 直接引用類別庫使用額外 .NET API |

### 2.3 運作環境
- Windows（COM 註冊需 `regasm.exe /codebase /tlb`，需系統管理員權限；x64／x86 呼叫端需分別註冊）。
- .NET Framework 4.7.2 以上（COM 用途）。
- 選配：LibreOffice（安裝版或免安裝版）。

### 2.4 設計限制與假設
1. 每個 `ExcelOperator` 物件同一時間只持有**一個**活頁簿（`_workbook`）與一個目前檔案路徑（`_filePath`）。
2. LibreOffice 設定與日誌設定存放於 **static** 的 `LibreOfficeConfigManager`，同一行程內所有 `ExcelOperator` 物件共用（見 ISS-13）。
3. 所有活頁簿操作在記憶體中進行，必須呼叫 `SaveWorkbook` 才會寫入磁碟（PDF／列印例外，見 FR-OUT-04）。
4. 支援格式以 ClosedXML 為準：`.xlsx`／`.xlsm` 等 OpenXML 格式；**不支援** 舊版 `.xls`。

---

## 3. 功能需求

> 欄位說明：**前置**＝呼叫前必要狀態；**行為**＝現況程式行為；**例外**＝會拋出的錯誤（COM 端表現為 `Err.Number <> 0`）。

### 3.1 版本資訊

#### FR-VER-01 取得版本資訊 `GetVersionInfo()`（DispId 7）
- 行為：回傳字串 `"{LibraryName} v{Version} ({LastUpdateDate})"`，目前為 `ExcelKit v1.2.4 (2026-04-05)`。
- 前置：無（不需活頁簿）。
- .NET 另提供靜態 `GetVersionInfoStatic()` 與公開靜態欄位 `Version`、`LibraryName`、`LastUpdateDate`。
- 備註：版本號同時寫在 csproj 與程式常數，需手動同步（ISS-17）。

### 3.2 活頁簿操作

#### FR-WB-01 建立新活頁簿 `CreateNewWorkbook()`（DispId 1）
- 行為：建立新的記憶體活頁簿，並自動新增工作表 `Sheet1`。
- 備註：不會釋放先前的活頁簿，也**不會清除** `_filePath`（ISS-04、ISS-05）。

#### FR-WB-02 開啟活頁簿 `OpenWorkbook(filePath)`（DispId 2）
- 行為：以 ClosedXML 讀入檔案，並將 `_filePath` 設為 `filePath`。
- 例外：檔案不存在 → `FileNotFoundException("檔案不存在: {filePath}")`；格式無效或檔案被鎖定 → ClosedXML／IO 例外原樣拋出。
- 備註：相對路徑以宿主行程的目前目錄解析。

#### FR-WB-03 儲存活頁簿 `SaveWorkbook(filePath)`（DispId 3）
- 行為：`filePath` 為空字串／null 時使用 `_filePath`（覆蓋原檔）；否則另存至 `filePath`，並將 `_filePath` 更新為新路徑。
- 例外：兩者皆空 → `InvalidOperationException("必須提供檔案路徑")`；尚未建立／開啟活頁簿 → `NullReferenceException`（ISS-01）；目的檔被其他程式鎖定 → IO 例外。
- 備註：COM 端參數不可省略，需傳 `""` 表示覆蓋原檔（手冊範例 `excel.SaveWorkbook ""`）。

### 3.3 工作表操作

#### FR-SH-01 新增工作表 `AddSheet(sheetName)`（DispId 14）
- 行為：新增指定名稱工作表，附加於最後。COM 端無回傳；.NET 端回傳 `IXLWorksheet`。
- 例外：名稱重複、超過 31 字元、含 `[ ] : * ? / \` 等非法字元、空白 → ClosedXML 例外。

#### FR-SH-02 刪除工作表 `DeleteSheet(sheetName)`（DispId 6）
- 行為：刪除指定名稱工作表。
- 例外：工作表不存在 → ClosedXML 例外。刪除最後一張工作表後存檔的行為未定義（ISS-08）。

#### FR-SH-03 取得所有工作表名稱 `GetAllSheetNames()`（DispId 26）
- 行為：依活頁簿順序回傳名稱。COM 端回傳 `string[]`（VBS 可 `For Each`）；.NET 端回傳 `List<string>`。

#### FR-SH-04 工作表自動建立（隱含行為）
- 行為：所有以 `sheetName` 為參數的儲存格與列欄方法（`SetCellValue*`、`GetCellValue*`、`ClearCell`、`InsertRow`、`DeleteRow`、`Set/GetRowHeight`、`InsertColumn`、`DeleteColumn`、`Set/GetColumnWidth`、`FormatCell`）在工作表不存在時會**自動新增**該工作表後再執行。
- 備註：讀取類操作也會產生副作用（新增空白工作表），且打錯名稱不會報錯（ISS-03）。`DeleteSheet` 不適用此行為。

### 3.4 儲存格操作

#### FR-CL-01 依 A1 位址寫入 `SetCellValue(sheetName, cellAddress, value)`（DispId 4）
- 行為：依 `value` 的 .NET 型別寫入，型別對應見 FR-CL-06。`value` 為 null（COM `Empty`／`Nothing`）時清除該儲存格。
- 例外：位址格式錯誤 → ClosedXML 例外。

#### FR-CL-02 依行列寫入 `SetCellValueByRC(sheetName, row, column, value)`（DispId 15）
- 行為：同 FR-CL-01，以 1-based 行列號定位（例：`2,3` = `C2`）。
- 例外：行或列 < 1、列 > 1,048,576、欄 > 16,384 → ClosedXML 例外。
- .NET 另有多載 `SetCellValue(sheetName, int row, int column, object value)`。

#### FR-CL-03 依 A1 位址讀取 `GetCellValue(sheetName, cellAddress)`（DispId 5）
- 行為（回傳型別對應）：

| 儲存格內容 | 回傳 |
| --- | --- |
| 空白 | `null`（COM：`Empty`） |
| 布林 | `bool` |
| 數值 | `double` |
| 日期時間 | `DateTime` |
| 錯誤值 | 錯誤代碼字串（如 `DivisionByZero`） |
| 文字及其他 | `string` |

- 備註：公式儲存格回傳 ClosedXML 計算後的值。

#### FR-CL-04 依行列讀取 `GetCellValueByRC(sheetName, row, column)`（DispId 16）
- 行為：同 FR-CL-03，以行列號定位。.NET 另有多載 `GetCellValue(sheetName, int, int)`。

#### FR-CL-05 清除儲存格 `ClearCell(sheetName, cellAddress)`（DispId 17）
- 行為：呼叫 ClosedXML `Cell.Clear()`（預設 `XLClearOptions.All`），清除**內容與格式**。
- 備註：手冊描述「保留格式」，與實際行為不符（ISS-10）。

#### FR-CL-06 寫入型別對應規則
| 傳入 .NET 型別 | 常見 COM／VBS 來源 | 寫入 Excel 型別 |
| --- | --- | --- |
| `null` | `Empty`、`Nothing` | 清除儲存格 |
| `string` | 字串 | 文字 |
| `int`（Int32） | VBS `Long`、LabVIEW I32 | 數值 |
| `double` | VBS `Double`、LabVIEW DBL | 數值 |
| `decimal` | VBS `Currency` | 數值（轉 double） |
| `bool` | VBS `Boolean` | 布林 |
| `DateTime` | VBS `Date` | 日期時間 |
| `long`（Int64） | — | 數值 |
| `float` | VBS `Single`、LabVIEW SGL | 數值 |
| 其他（含 `short`、`byte`） | **VBS `Integer`（如字面值 `123`）**、`Byte`、LabVIEW I16／U8 | **以 `ToString()` 寫成文字** |

- 備註：VBScript 中小於 32768 的整數字面值為 `Integer`（Int16），目前會被存成文字（ISS-02）。

#### FR-CL-07 儲存格格式（僅 .NET）`FormatCell(sheetName, cellAddress, Action<IXLStyle>)`
- 行為：對儲存格樣式執行委派。未列入 COM 介面，COM Client 無法使用。

### 3.5 列與欄操作

| 編號 | 方法（DispId） | 行為 |
| --- | --- | --- |
| FR-RC-01 | `InsertRow(sheetName, rowNumber)`（18） | 在 `rowNumber` **上方**插入一列，原列以下下移 |
| FR-RC-02 | `DeleteRow(sheetName, rowNumber)`（19） | 刪除該列，以下各列上移 |
| FR-RC-03 | `SetRowHeight(sheetName, rowNumber, height)`（20）／`GetRowHeight`（21） | 設定／取得列高（點） |
| FR-RC-04 | `InsertColumn(sheetName, columnNumber)`（22） | 在 `columnNumber` **左方**插入一欄，原欄以右右移 |
| FR-RC-05 | `DeleteColumn(sheetName, columnNumber)`（23） | 刪除該欄，以右各欄左移 |
| FR-RC-06 | `SetColumnWidth(sheetName, columnNumber, width)`（24）／`GetColumnWidth`（25） | 設定／取得欄寬（字元寬度單位） |

- 共通：行列號 1-based；超出範圍或 ≤0 → ClosedXML 例外；工作表不存在時自動建立（FR-SH-04）。
- 備註：手冊 5.3 使用「列／行」用語與程式註解（「行操作」＝Row）不一致，閱讀時以英文方法名稱為準。

### 3.6 PDF 匯出與列印（需 LibreOffice）

#### FR-OUT-01 匯出 PDF `ExportToPdf(excelFilePath, pdfOutputPath)`（DispId 8）
處理流程：
1. 記憶體同步（FR-OUT-04）。
2. 檢查來源檔存在，否則 `FileNotFoundException`。
3. 兩個路徑轉為絕對路徑。
4. 取得 LibreOffice 路徑（FR-OUT-03），取不到 → 例外（訊息含「LibreOffice 路徑未設定或不存在」）。
5. 執行 `soffice --headless --convert-to pdf --outdir "<輸出資料夾>" "<來源檔>"`。輸出資料夾為磁碟根目錄時改用 `X:\.`，其他情況去除尾端 `\`。
6. 逾時（30 秒）→ 終止行程並拋 `TimeoutException`；結束碼 ≠ 0 → 例外並附 stderr。
7. LibreOffice 產出 `<來源主檔名>.pdf`，若與 `pdfOutputPath` 不同則**刪除既有目的檔**後搬移更名。
8. 寫入 INFO 日誌；任何錯誤寫 ERROR 日誌後統一包裝為 `Exception("PDF 匯出失敗：...")`。
- 備註：LibreOffice 結束碼為 0 但未產出 PDF 時不會報錯（ISS-09）。

#### FR-OUT-02 列印 `PrintWorkbook(excelFilePath, printerName)`（DispId 9）
- 行為：記憶體同步（FR-OUT-04）後，`printerName` 空白時執行 `soffice --headless -p "<檔案>"`（系統預設印表機），否則 `soffice --headless --pt "<印表機>" "<檔案>"`。逾時 30 秒、結束碼檢查、日誌與錯誤包裝同 FR-OUT-01，錯誤訊息前綴「列印失敗:」。
- 備註：結束碼 0 僅代表命令送出，不代表實體列印成功。.NET 另有 `Print(excelFilePath, printerName = null)`。

#### FR-OUT-03 LibreOffice 路徑解析順序
1. `SetLibreOfficePath` 指定且檔案存在。
2. 系統 PATH（Windows `where soffice`，其他平台 `which soffice`）。
3. `C:\Program Files\LibreOffice\program\soffice.exe`。
4. `C:\Program Files (x86)\LibreOffice\program\soffice.exe`。
5. 以上皆無 → 回傳 null（呼叫端拋錯）。

#### FR-OUT-04 匯出／列印前自動存檔
- 行為：若物件目前持有活頁簿，先以 `SaveWorkbook(目標路徑)` 存檔再交給 LibreOffice。目標路徑 = `excelFilePath`（非空白時）否則 `_filePath`。
- 影響：**傳入的 `excelFilePath` 會被目前記憶體活頁簿覆寫**，即使該路徑原本是另一個檔案（ISS-06）；同時 `_filePath` 會被改為該路徑。

#### FR-OUT-05 逾時
- 固定 30 秒（`GetTimeoutSeconds()` 回傳常數），不可設定。

#### FR-OUT-06 匯出 PDF（僅 .NET 多載）`ExportToPdf(excelFilePath, pdfOutputPath, libreOfficePath)`
- 行為：使用參數指定的 soffice，不讀取設定、不寫日誌；等待 30 秒但未判斷是否逾時（ISS-11）。

### 3.7 設定與日誌

| 編號 | 方法（DispId） | 行為 |
| --- | --- | --- |
| FR-CFG-01 | `SetLibreOfficePath(path)`（10） | 設定 soffice 路徑覆寫值；空白則清除覆寫值。設定當下不驗證檔案存在，於使用時才檢查 |
| FR-CFG-02 | `SetLogDirectory(path)`（11） | 設定日誌資料夾；空白則恢復預設（DLL 所在資料夾）。資料夾不存在時於寫入時自動建立 |
| FR-CFG-03 | `SetLogEnabled(enabled)`（12） | 啟用／停用日誌，預設啟用 |
| FR-CFG-04 | `SetLogLevel(level)`（13） | 1=ERROR、2=WARN、3=INFO（預設 3）。<1 夾為 1，>3 夾為 3 |

#### FR-CFG-05 日誌格式
- 檔名：`{組件名稱}{組件版本}_{yyyyMMdd}.txt`，例：`ExcelKit1.2.4.0_20261114.txt`。
- 行格式：`[yyyy-MM-dd HH:mm:ss] [ERROR|WARN|INFO] 訊息`，UTF-8 附加寫入。
- 記錄條件：已啟用且「設定等級 ≥ 訊息等級」。
- 記錄事件：PDF 匯出開始／成功／失敗、列印開始／完成／失敗。活頁簿與儲存格操作不記錄。
- 寫入失敗一律忽略，不影響主流程。
- 範圍：設定為行程內共用（static），影響同一行程所有物件。

### 3.8 資源釋放
- FR-RES-01 `Dispose()`：釋放活頁簿。未列入 COM 介面；COM Client 以 `Set obj = Nothing` 釋放，實際回收依 .NET GC 時機。

---

## 4. 外部介面需求

### 4.1 COM 介面總表（`IExcelOperator`，IDispatch）

| DispId | 方法 | 參數 | 回傳 | 需求 |
| --- | --- | --- | --- | --- |
| 1 | CreateNewWorkbook | — | — | FR-WB-01 |
| 2 | OpenWorkbook | filePath: String | — | FR-WB-02 |
| 3 | SaveWorkbook | filePath: String | — | FR-WB-03 |
| 4 | SetCellValue | sheetName, cellAddress: String; value: Variant | — | FR-CL-01 |
| 5 | GetCellValue | sheetName, cellAddress: String | Variant | FR-CL-03 |
| 6 | DeleteSheet | sheetName: String | — | FR-SH-02 |
| 7 | GetVersionInfo | — | String | FR-VER-01 |
| 8 | ExportToPdf | excelFilePath, pdfOutputPath: String | — | FR-OUT-01 |
| 9 | PrintWorkbook | excelFilePath, printerName: String | — | FR-OUT-02 |
| 10 | SetLibreOfficePath | libreOfficePath: String | — | FR-CFG-01 |
| 11 | SetLogDirectory | logDirectoryPath: String | — | FR-CFG-02 |
| 12 | SetLogEnabled | enabled: Boolean | — | FR-CFG-03 |
| 13 | SetLogLevel | logLevel: Long | — | FR-CFG-04 |
| 14 | AddSheet | sheetName: String | — | FR-SH-01 |
| 15 | SetCellValueByRC | sheetName: String; row, column: Long; value: Variant | — | FR-CL-02 |
| 16 | GetCellValueByRC | sheetName: String; row, column: Long | Variant | FR-CL-04 |
| 17 | ClearCell | sheetName, cellAddress: String | — | FR-CL-05 |
| 18 | InsertRow | sheetName: String; rowNumber: Long | — | FR-RC-01 |
| 19 | DeleteRow | sheetName: String; rowNumber: Long | — | FR-RC-02 |
| 20 | SetRowHeight | sheetName: String; rowNumber: Long; height: Double | — | FR-RC-03 |
| 21 | GetRowHeight | sheetName: String; rowNumber: Long | Double | FR-RC-03 |
| 22 | InsertColumn | sheetName: String; columnNumber: Long | — | FR-RC-04 |
| 23 | DeleteColumn | sheetName: String; columnNumber: Long | — | FR-RC-05 |
| 24 | SetColumnWidth | sheetName: String; columnNumber: Long; width: Double | — | FR-RC-06 |
| 25 | GetColumnWidth | sheetName: String; columnNumber: Long | Double | FR-RC-06 |
| 26 | GetAllSheetNames | — | Variant（String 陣列） | FR-SH-03 |

- DispId 為 COM 相容契約，**新增方法只能追加新 DispId，不得變更既有編號或簽章**（LabVIEW VI 以介面綁定）。
- 手冊 6 章「常數」`LogLevelError` 等為 .NET `public const`，COM 端不可存取，需直接使用數值 1～3。

### 4.2 錯誤回報介面
- 所有錯誤以 .NET 例外拋出，COM 端取得 `Err.Number`（HRESULT）與 `Err.Description`（例外訊息）。
- PDF／列印錯誤統一包裝為 `System.Exception`（HRESULT `0x80131500`），原始例外型別不保留；呼叫端須以訊息文字判斷（測試腳本以訊息含 `libreoffice` 判斷為略過）。

### 4.3 外部程式介面（LibreOffice CLI）
| 用途 | 命令列 |
| --- | --- |
| PDF | `soffice --headless --convert-to pdf --outdir "<dir>" "<xlsx>"` |
| 預設印表機 | `soffice --headless -p "<xlsx>"` |
| 指定印表機 | `soffice --headless --pt "<printer>" "<xlsx>"` |

### 4.4 檔案介面
| 檔案 | 讀／寫 | 說明 |
| --- | --- | --- |
| 活頁簿 `.xlsx` | 讀寫 | ClosedXML |
| PDF | 寫 | 由 LibreOffice 產出後更名，會覆寫同名目的檔 |
| 日誌 `.txt` | 附加寫 | FR-CFG-05 |

---

## 5. 非功能需求

| 編號 | 類別 | 需求（現況） | 驗證方式 |
| --- | --- | --- | --- |
| NFR-01 | 相容性 | COM 用途需 .NET Framework 4.7.2；x86 與 x64 呼叫端需以對應 regasm 註冊 | 兩種位元各執行 TC-ENV-01 |
| NFR-02 | 相容性 | net10.0 組建僅作 .NET 函式庫使用；未啟用 COM Hosting，無法作為 COM 元件（ISS-16） | 建置檢查 |
| NFR-03 | 部署 | net472 以 Costura 將相依 DLL 內嵌，部署單一 `ExcelKit.dll`（＋`.tlb`） | 乾淨環境註冊測試 |
| NFR-04 | 效能 | PDF 匯出／列印單次逾時 30 秒；其他操作為記憶體操作，無明訂效能指標（未量測，待定義） | TC-OUT-07、TC-PERF-01 |
| NFR-05 | 可靠性 | 外部行程逾時須被終止，不得殘留阻塞呼叫端 | TC-OUT-07 |
| NFR-06 | 可診斷性 | PDF／列印流程具分級日誌；日誌失敗不影響主流程 | TC-CFG-* |
| NFR-07 | 執行緒 | 未設計執行緒安全；同一物件不得多執行緒同時呼叫；設定為行程內共用 | 文件約束 |
| NFR-08 | 資料安全 | 匯出時會覆寫來源 Excel 與目的 PDF，呼叫端須確保路徑正確（ISS-06） | TC-OUT-05、TC-OUT-06 |
| NFR-09 | 可維護性 | 版本號需於 csproj 與 `ExcelOperator.Version`／`LastUpdateDate` 同步 | TC-VER-02 |
| NFR-10 | 授權 | 散布時需遵循 ClosedXML 等 MIT 授權及 LibreOffice 授權 | 發布檢查 |

---

## 6. 五項適用性判定（本文件為現況整理）

| 面向 | 判定 | 理由 |
| --- | --- | --- |
| 架構 | 不適用 | 本次僅整理文件，不改程式結構；文件中記錄的架構風險（static 設定、COM GUID）由後續 Issue 評估 |
| UI／視覺 | 不適用 | 元件無使用者介面 |
| 資料庫 | 不適用 | 無應用資料庫 |
| 系統整合 | 待確認 | 元件本身為整合介面（COM、LabVIEW、LibreOffice）；任何介面修正都需系統整合評估 |
| 維運 | 待確認 | 日誌預設寫在 DLL 資料夾、需管理員註冊；固定發布資料夾與正式版基準尚未確認 |

## 7. 範圍外與待決事項
1. 發布資料夾、正式版基準與 COM 註冊流程是否納入安裝程式：未確認。
2. 問題清單 `ISS-xx` 是否修正、修正後是否需新版號：待人類決策。
3. 效能指標（大量資料寫入、大檔開啟）：無既有需求，未定義。

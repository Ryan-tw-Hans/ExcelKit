# 程式碼問題與風險清單｜ExcelKit v1.2.4

| 項目 | 內容 |
| --- | --- |
| 文件編號 | CR-ExcelKit-v1.2.4 |
| 文件版本 | v0.1（草稿） |
| 日期 | 2026-11-14（Asia/Taipei） |
| 執行者 | Claude（SA／測試工程師角色，靜態閱讀程式碼；自查，非獨立審查） |
| 審查範圍 | `Excelopen/Class1.cs`（978 行）、`Excelopen.csproj`、`AssemblyInfo.cs`、`測試腳本/*.vbs`、使用手冊 |
| 方法 | 靜態閱讀與推論。**未建置、未執行**，下列「影響」皆為依程式碼推得，需以 TCS 對應案例實測確認 |

嚴重度：**高**＝可能造成資料遺失、錯誤資料或與其他系統衝突；**中**＝功能錯誤或難以診斷；**低**＝一致性、可維護性。

## 1. 問題總表

| 編號 | 嚴重度 | 分類 | 摘要 | 位置 | 驗證案例 |
| --- | --- | --- | --- | --- | --- |
| ISS-06 | 高 | 資料安全 | 匯出 PDF／列印會用記憶體活頁簿覆寫 `excelFilePath` 指向的檔案 | `ExportToPdf`、`Print` 開頭同步區塊 | TC-OUT-05 |
| ISS-07 | 高 | COM 整合 | COM 介面／類別 GUID 為佔位值 `12345678-…`、`87654321-…`，易與其他元件衝突 | `Class1.cs` L16、L80 | TC-ENV-03 |
| ISS-02 | 高 | 資料正確性 | VBS `Integer`（Int16）、`Byte` 等型別被存成文字 | `SetCellValue` 型別判斷 | TC-CL-08 |
| ISS-05 | 高 | 資料安全 | `CreateNewWorkbook` 不清除 `_filePath`，之後 `SaveWorkbook ""` 會覆寫先前開啟的檔案 | `CreateNewWorkbook` | TC-WB-09 |
| ISS-01 | 中 | 錯誤處理 | 未建立／開啟活頁簿就呼叫操作 → `NullReferenceException`，訊息不具意義 | 所有使用 `_workbook` 的方法 | TC-WB-08 |
| ISS-03 | 中 | 行為 | 工作表不存在時自動新增，連讀取都會新增；名稱打錯不報錯 | `GetSheet` | TC-SH-06、TC-SH-07 |
| ISS-09 | 中 | 錯誤處理 | LibreOffice 結束碼 0 但未產出 PDF 時仍視為成功 | `ExportToPdf` 更名區塊 | TC-OUT-08 |
| ISS-12 | 中 | 可靠性 | 重導 stdout／stderr 但在 `WaitForExit` 前未讀取，輸出量大時可能死結至逾時 | `ExportToPdf`、`Print`、`TryFindInSystemPath` | 程式審查 |
| ISS-13 | 中 | 架構 | 設定管理器為 static 且非執行緒安全；一個物件改設定會影響同行程所有物件 | `_configManager` | TC-CFG-06 |
| ISS-15 | 中 | 整合 | 已有 LibreOffice 視窗或另一個 headless 行程使用同一使用者設定檔時，轉換可能靜默失敗或卡住 | LibreOffice 呼叫 | TC-OUT-10 |
| ISS-19 | 中 | 測試 | 既有 VBS 測試有互動視窗、硬編碼、缺少驗證等問題，無法無人值守 | `測試腳本/` | 見第 3 節 |
| ISS-04 | 低 | 資源 | 重新 Create／Open 時未 Dispose 舊活頁簿 | `CreateNewWorkbook`、`OpenWorkbook` | TC-PERF-02 |
| ISS-08 | 低 | 行為 | 可刪除最後一張工作表，存檔結果未定義 | `DeleteSheet` | TC-SH-05 |
| ISS-10 | 低 | 文件一致性 | `ClearCell` 實際清除內容＋格式，手冊寫「保留格式」 | `ClearCell`、手冊 5.2 | TC-CL-07 |
| ISS-11 | 低 | 可靠性 | .NET 三參數 `ExportToPdf` 忽略 `WaitForExit` 結果，逾時時讀 `ExitCode` 會再拋例外且不終止行程 | L661 | 程式審查（非 COM） |
| ISS-14 | 低 | 錯誤處理 | PDF／列印錯誤一律包成 `System.Exception`，丟失例外型別與堆疊 | catch 區塊 | TC-OUT-03 |
| ISS-16 | 低 | 建置 | net10.0 目標未設 `EnableComHosting`，不可作 COM 元件；README 未說明 | csproj | 建置檢查 |
| ISS-17 | 低 | 維護 | 版本號分散於 csproj 與程式常數；`LastUpdateDate` 手動維護 | csproj、L92–102 | TC-VER-02 |
| ISS-18 | 低 | 維運 | 預設日誌寫在 DLL 資料夾，若安裝於 Program Files 將無寫入權限且被靜默忽略 | `GetLogDirectoryPath` | TC-CFG-05 |
| ISS-20 | 低 | 文件一致性 | 命名空間 `Excelopen` 與組件名 `ExcelKit` 不一致；手冊列的 `LogLevelError` 等常數 COM 端無法存取 | `Class1.cs` L10、手冊 6 章 | 文件審查 |

## 2. 問題說明與建議

### ISS-06（高）匯出／列印會覆寫來源檔
`ExportToPdf`／`Print` 開頭：只要 `_workbook != null`，就執行 `SaveWorkbook(excelFilePath 或 _filePath)`。
- 情境：`OpenWorkbook "A.xlsx"` 後呼叫 `ExportToPdf "B.xlsx", "B.pdf"` → `B.xlsx` 被 A 的內容覆寫，產出的 PDF 也是 A。
- 情境：只想把磁碟上的既有檔轉 PDF，但物件內仍留有先前的活頁簿 → 同樣被覆寫。
- 建議：僅在 `excelFilePath` 與 `_filePath` 指向同一檔（或為空）時才同步存檔；否則直接轉換磁碟檔。或增加明確的「先存再轉」選項。

### ISS-07（高）佔位 GUID
`IExcelOperator` 用 `12345678-1234-1234-1234-123456789012`、`ExcelOperator` 用 `87654321-4321-4321-4321-210987654321`，屬常見範例值。若另一個元件也照範例使用，註冊後會互相覆蓋登錄 `HKCR\CLSID`／`Interface`，造成 CreateObject 載入錯誤元件。
- 建議：產生正式 GUID。**注意此為破壞性變更**，LabVIEW VI 需重新綁定、須重新註冊，屬 L3（跨系統介面）並需人類核准。

### ISS-02（高）VBS Integer 被存成文字
VBScript 數字字面值在 -32768～32767 為 `Integer`（VT_I2 → `short`）；程式未處理 `short`、`byte`、`ushort`、`uint`、`sbyte`、`ulong`，落入 `ToString()` 分支。結果 `SetCellValue "Sheet1","A1",123` 在 Excel 內為文字「123」，公式加總失效、出現綠色三角形警告。LabVIEW I16／U8／U16／U32 同樣受影響。
- 建議：以 `IConvertible` 判斷所有數值型別一律轉 `double`。兩個 `SetCellValue` 多載重複邏輯可合併為共用的 `AssignValue(IXLCell, object)`。

### ISS-05（高）新建活頁簿保留舊路徑
`OpenWorkbook "Template.xlsx"` → `CreateNewWorkbook` → `SaveWorkbook ""` → 空白活頁簿覆寫 `Template.xlsx`。
- 建議：`CreateNewWorkbook` 時 `_filePath = null`。

### ISS-01（中）未初始化狀態
所有操作直接使用 `_workbook`，未檢查 null。COM 端只看到「並未將物件參考設定為物件的執行個體」。
- 建議：加 `EnsureWorkbook()`，拋 `InvalidOperationException("請先呼叫 CreateNewWorkbook 或 OpenWorkbook")`。

### ISS-03（中）工作表自動建立
`GetSheet` 以 try/catch 吞掉「找不到工作表」後新增。`GetCellValue "Shee1","A1"`（打錯字）回傳 Empty 並多一張空白工作表，存檔後才被發現。另 catch 範圍過大，其他錯誤也會觸發新增。
- 建議：讀取類操作不自動建立，改拋明確錯誤；寫入類是否自動建立由人類決定（維持相容可加開關）。改用 `TryGetWorksheet`。

### ISS-09（中）未驗證 PDF 產出
LibreOffice 遇到檔案鎖定、設定檔衝突等情況常回傳 0 卻沒產出檔案。目前 `File.Exists(actualPdfPath)` 為 false 時直接略過並記 INFO「成功」。
- 建議：轉換後確認 `pdfOutputPath` 存在且大小 > 0，否則拋錯。

### ISS-12（中）標準輸出緩衝死結
`RedirectStandardOutput/Error = true`，但僅在結束後才 `ReadToEnd()`。子行程輸出超過管線緩衝（約 4 KB）時會阻塞，主程式等到 30 秒逾時才失敗。`TryFindInSystemPath` 的 `WaitForExit()` 無逾時。
- 建議：使用 `BeginOutputReadLine/BeginErrorReadLine` 非同步讀取，或不重導 stdout。

### ISS-13（中）static 設定
`_configManager` 為 static、延遲建立無鎖定。LabVIEW 可能在同一行程中平行呼叫多個物件，設定互相影響，且有競態條件。
- 建議：至少以 `Lazy<T>` 建立並在手冊中說明為「行程層級設定」；若需物件層級設定則改為實例欄位。

### ISS-15（中）LibreOffice 執行個體衝突
未指定 `-env:UserInstallation`，與使用者已開啟的 LibreOffice 共用設定檔，headless 呼叫可能被轉交給既有行程後立即以 0 結束，或卡住。既有測試報告有「另一個處理序正在使用檔案」紀錄。
- 建議：加入 `-env:UserInstallation=file:///<temp 專屬資料夾>`，並在手冊列為已知限制。

### ISS-19（中）測試腳本問題
見第 3 節。

### 低嚴重度建議（摘要）
- ISS-04：Create／Open 前先 `_workbook?.Dispose()`。
- ISS-08：刪除前檢查剩餘工作表數 > 1。
- ISS-10：改用 `Clear(XLClearOptions.Contents)` 或修正手冊（需決定何者為準）。
- ISS-11：比照兩參數版本處理逾時，或移除此多載改呼叫共用私有方法（目前兩個 ExportToPdf 有大量重複程式碼）。
- ISS-14：自訂例外或保留 `InnerException`，並定義錯誤代碼表供 LabVIEW 判斷。
- ISS-16：README 註明 net10 用途；若要 COM 則加 `EnableComHosting`。
- ISS-17：程式版本改由組件屬性讀取（`AssemblyInformationalVersion`），避免手動同步。
- ISS-18：預設日誌改至 `%ProgramData%\ExcelKit\Logs` 或 `%TEMP%`。
- ISS-20：手冊第 6 章註明常數需以數值傳入。

## 3. 既有測試腳本問題（ISS-19 細項）

| 細項 | 腳本 | 問題 | 影響 |
| --- | --- | --- | --- |
| T-01 | `05_Set_Cell_By_Name.vbs` | 殘留 `msgbox value` | 無人值守執行會卡住 |
| T-02 | `10_Print_WorkBook.vbs` | 印表機名硬編碼為 `"Test.xlsx"`，且有兩個 `msgbox` | 實際送到不存在的印表機名稱；卡住 |
| T-03 | `_TestCommon.vbs` `GetDllPath` | 預設 DLL 名為 `Excelopen.dll`，實際輸出為 `ExcelKit.dll` | 報告顯示 `DLL Version: UnknownVersion`（已在 2026-08-15 報告中出現） |
| T-04 | `04`、`12`、`13`～`19` | `LogTestResult` 測試名稱傳空字串（顯示名稱由編號帶出，功能不受影響） | 可讀性 |
| T-05 | `12_Delete_Sheet.vbs` | 只確認無錯誤，未確認工作表已消失 | 驗證不足 |
| T-06 | `19_Get_All_Sheet_Names.vbs` | 未確認回傳內容為 `Sheet1, Data` | 驗證不足 |
| T-07 | `20_Config_Functions.vbs` | 未確認日誌檔產生與內容 | 驗證不足 |
| T-08 | `17_Insert_Delete_Column.vbs` | `On Error Resume Next` 下先比對值再檢查 `Err`，中途錯誤訊息可能被覆蓋 | 失敗原因不準確 |
| T-09 | `05`、`06` | 以 `Microsoft.ACE.OLEDB.12.0` 驗證，需另裝 Access Database Engine 且位元需一致 | 環境相依，可能誤判失敗 |
| T-10 | `Run_All_Ordered_Tests.vbs` | 以資料夾列舉順序執行，Windows 不保證排序 | 執行順序不穩定 |
| T-11 | `TotalTest.vbs` | 順序為 03→08→04，與編號不同；需確認變數宣告完整 | 可讀性 |
| T-12 | 全部 | 無負面測試（不存在檔案、非法名稱、越界）、無型別測試 | 覆蓋缺口，見 RTM |
| T-13 | `_debug*.vbs` | 除錯殘留檔 | 建議清理（需人類確認） |

## 4. 建議處理順序
1. **先決策**：ISS-06、ISS-05、ISS-02（資料正確性，修改小、價值高）→ 可合為一個 L2 Issue。
2. ISS-07 GUID 變更屬 L3，需 SA 評估與 LabVIEW 端相容計畫。
3. ISS-19 測試腳本整理（移除 msgbox、修正 DLL 名、補驗證）→ L1／L2，可先做以建立回歸基線。
4. 其餘中低嚴重度項目排入後續版本。

所有修正均需依 `AGENTS.md` 由 PM 受理、判級並寫入 IssueLog，本文件不代表已核准修正。

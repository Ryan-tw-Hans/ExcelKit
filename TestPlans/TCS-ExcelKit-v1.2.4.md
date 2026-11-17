# 測試案例規格（TCS）｜ExcelKit v1.2.4

| 項目 | 內容 |
| --- | --- |
| 文件編號 | TCS-ExcelKit-v1.2.4 |
| 文件版本 | v0.1（草稿） |
| 日期 | 2026-11-14（Asia/Taipei） |
| 撰寫者 | Claude（測試工程師角色；自查，非獨立審查） |
| 依據 | SRS-ExcelKit-v1.2.4、CR-ExcelKit-v1.2.4、STP-ExcelKit-v1.2.4 |

## 0. 共通約定
- `op`：`CreateObject("ExcelKit.ExcelOperator")`；`op2`：另一個新物件，用於讀回驗證。
- `WB`：`測試腳本\TestData\TC_<案例編號>.xlsx`；`PDF`：同資料夾 `.pdf`。案例開始前刪除舊檔。
- 「錯誤」＝ VBS `On Error Resume Next` 後 `Err.Number <> 0`。
- 「既有腳本」欄：可直接沿用或需補強的 `測試腳本/` 檔案；「新增」表示需新寫。
- 標示 **ISS-xx** 者依 STP 3.1 判定：先列「現況預期」，再列「期望行為（修正後）」。
- 全部案例狀態：**未執行**。

---

## 1. 建置與環境（TC-BLD／TC-ENV）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-BLD-01 | 雙目標建置 | NFR-01、02 | P1 | `dotnet build Excelopen.slnx -c Release` | 無錯誤；產生 `bin\Release\net472\ExcelKit.dll` 與 `net10.0\ExcelKit.dll` | 新增 |
| TC-BLD-02 | net472 單一 DLL | NFR-03 | P2 | 將 net472 的 `ExcelKit.dll` 單獨複製到乾淨資料夾後註冊並執行 TC-ENV-01 | 不缺 ClosedXML 等相依即可建立物件與寫檔 | 新增 |
| TC-BLD-03 | 型別庫產生 | 4.1 | P2 | `regasm /codebase ExcelKit.dll /tlb ExcelKit.tlb` | 產生 `.tlb`，OLE/COM Viewer 可見 26 個方法，DispId 與 SRS 4.1 相同 | 新增 |
| TC-ENV-01 | x64 建立 COM 物件 | NFR-01 | P1 | x64 註冊後以 `System32\cscript` 建立 `op` | 無錯誤，物件非 Nothing | `01_Reference_COM_Component.vbs` |
| TC-ENV-02 | x86 建立 COM 物件 | NFR-01 | P1 | x86 註冊後以 `SysWOW64\cscript` 執行 TC-ENV-01、TC-WB-02 | 皆通過 | 新增（以既有腳本切換 cscript） |
| TC-ENV-03 | GUID 衝突檢查 **ISS-07** | 4.1 | P2 | 檢查 `HKCR\CLSID\{87654321-4321-4321-4321-210987654321}` 與 `HKCR\Interface\{12345678-1234-1234-1234-123456789012}` 的 InprocServer32／CodeBase 與 TypeLib | 現況預期：指向 ExcelKit.dll；若指向其他元件即為衝突。期望行為：改用正式 GUID | 新增（手動） |
| TC-ENV-04 | 未註冊時的錯誤 | NFR-01 | P3 | 解除註冊後建立物件 | 錯誤 429（ActiveX 元件無法建立物件） | 新增 |

## 2. 版本（TC-VER）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-VER-01 | 版本字串格式 | FR-VER-01 | P1 | `v = op.GetVersionInfo()`，不建立活頁簿 | `v = "ExcelKit v1.2.4 (2026-04-05)"` | `02_Read_DLL_Version.vbs`（需補字串比對） |
| TC-VER-02 | 版本一致性 | NFR-09 | P2 | 比對 `GetVersionInfo`、DLL 檔案版本（`fso.GetFileVersion`）、csproj `Version` | 三者主版本號一致：1.2.4／1.2.4.0 | 新增（修正 T-03 DLL 名後） |

## 3. 活頁簿（TC-WB）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-WB-01 | 新建含 Sheet1 | FR-WB-01 | P1 | `CreateNewWorkbook`；`GetAllSheetNames` | 陣列僅 `Sheet1` | `03_Create_WorkBook.vbs`（補名稱驗證） |
| TC-WB-02 | 另存新檔 | FR-WB-03 | P1 | Create；A1="SaveTest_<亂數>"；`SaveWorkbook WB`；`op2.OpenWorkbook WB`；讀 A1 | 檔案存在；讀回值相同 | `08_Save_WorkBook.vbs` |
| TC-WB-03 | 開啟既有檔 | FR-WB-02 | P1 | 先產生 WB；`op2.OpenWorkbook WB`；讀 A1 | 無錯誤；值相同 | `11_Open_WorkBook.vbs` |
| TC-WB-04 | 以空字串覆寫原檔 | FR-WB-03 | P1 | Open WB；A1 改為新值；`SaveWorkbook ""`；op2 開啟讀回 | 原檔被更新為新值 | 新增 |
| TC-WB-05 | 無路徑儲存 | FR-WB-03 | P2 | Create；`SaveWorkbook ""` | 錯誤，訊息含「必須提供檔案路徑」 | 新增 |
| TC-WB-06 | 開啟不存在檔 | FR-WB-02 | P2 | `OpenWorkbook "…\NotExist.xlsx"` | 錯誤，訊息含「檔案不存在」 | 新增 |
| TC-WB-07 | 開啟非 xlsx | 2.4-4 | P2 | 分別開啟 `.txt`（改副檔名為 .xlsx）與舊版 `.xls` | 錯誤；物件不崩潰，之後 Create 可正常使用 | 新增 |
| TC-WB-08 | 未初始化即操作 **ISS-01** | FR-WB-01 | P2 | 新物件直接 `SetCellValue "Sheet1","A1","x"`、`SaveWorkbook WB` | 現況預期：錯誤，訊息為 NullReference（並未將物件參考設定…）。期望行為：錯誤訊息提示先建立／開啟活頁簿 | 新增 |
| TC-WB-09 | 新建後覆寫舊路徑 **ISS-05** | FR-WB-01、03 | P1 | 產生 WB（A1="Original"）；`OpenWorkbook WB`；`CreateNewWorkbook`；`SaveWorkbook ""`；op2 開啟 WB 讀 A1 | 現況預期：無錯誤，A1 為 Empty（原檔被空白活頁簿覆寫）。期望行為：`SaveWorkbook ""` 回報需提供路徑，原檔 A1 仍為 "Original" | 新增 |
| TC-WB-10 | 儲存至被鎖定檔 | FR-WB-03 | P2 | 以 Excel 開啟 WB 後 `SaveWorkbook WB` | 錯誤，訊息指出檔案被使用；物件仍可另存其他路徑 | 新增（手動） |
| TC-WB-11 | 中文與空白路徑 | FR-WB-02、03 | P2 | 路徑 `TestData\測試 資料夾\報表 1.xlsx` 存檔再開啟 | 成功讀回 | 新增 |

## 4. 工作表（TC-SH）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-SH-01 | 新增工作表 | FR-SH-01 | P1 | Create；`AddSheet "Data"`；存檔；op2 開啟取名稱 | `Sheet1, Data`（順序相同） | `04_Create_WorkSheet.vbs`（補驗證） |
| TC-SH-02 | 重複名稱 | FR-SH-01 | P2 | Create；`AddSheet "Sheet1"` | 錯誤 | 新增 |
| TC-SH-03 | 非法名稱 | FR-SH-01 | P2 | 分別 `AddSheet` 32 字元名稱、`"A/B"`、`"A[1]"`、`""` | 皆錯誤；活頁簿仍只有 Sheet1 | 新增 |
| TC-SH-04 | 刪除工作表 | FR-SH-02 | P1 | Create；AddSheet "ToDelete"；DeleteSheet "ToDelete"；GetAllSheetNames | 僅 `Sheet1` | `12_Delete_Sheet.vbs`（補驗證，T-05） |
| TC-SH-05 | 刪除最後一張 **ISS-08** | FR-SH-02 | P3 | Create；DeleteSheet "Sheet1"；SaveWorkbook WB；op2 開啟 | 現況預期：記錄實際結果（刪除是否成功、存檔／開啟是否錯誤）。期望行為：刪除最後一張時回報錯誤 | 新增 |
| TC-SH-06 | 讀取不存在工作表 **ISS-03** | FR-SH-04 | P2 | Create；`v = GetCellValue("NoSuch","A1")`；GetAllSheetNames | 現況預期：無錯誤，v 為 Empty，名稱多出 `NoSuch`。期望行為：錯誤，不新增工作表 | 新增 |
| TC-SH-07 | 寫入不存在工作表 | FR-SH-04 | P2 | Create；`SetCellValue "Auto","A1","x"`；存檔讀回 | 現況：自動新增 `Auto` 且 A1="x"（是否保留此行為待決策） | 新增 |
| TC-SH-08 | 刪除不存在工作表 | FR-SH-02 | P2 | Create；`DeleteSheet "NoSuch"` | 錯誤；Sheet1 不受影響 | 新增 |
| TC-SH-09 | 列舉名稱 | FR-SH-03 | P1 | Create；AddSheet "B"、"A"；`names = GetAllSheetNames()`；`For Each` | `UBound=2`，依序 `Sheet1, B, A` | `19_Get_All_Sheet_Names.vbs`（補驗證，T-06） |

## 5. 儲存格（TC-CL）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-CL-01 | A1 寫入字串 | FR-CL-01 | P1 | Create；A1=動態字串；存檔；op2 讀 A1 | 值相同 | `05_Set_Cell_By_Name.vbs`（移除 msgbox，T-01） |
| TC-CL-02 | RC 寫入 | FR-CL-02 | P1 | `SetCellValueByRC "Sheet1",2,3,s`；存檔；op2 `GetCellValue("Sheet1","C2")` | 值相同（確認 2,3 = C2） | `06_Set_Cell_By_Row_Column.vbs` |
| TC-CL-03 | A1 讀取 | FR-CL-03 | P1 | B2 寫入後存檔，op2 開啟讀 B2 | 值相同，`TypeName = "String"` | `07_Read_Cell_Value.vbs` |
| TC-CL-04 | RC 讀取 | FR-CL-04 | P1 | 寫 C2 後 `GetCellValueByRC("Sheet1",2,3)` | 值相同 | 新增 |
| TC-CL-05 | 型別往返 | FR-CL-03、06 | P1 | 寫入 `CDbl(3.14)`、`CLng(100000)`、`True`、`CDate("2026-11-14 10:30")`、`CCur(12.5)`、`CSng(1.5)`；存檔；op2 讀回並取 `TypeName` | 依序 Double 3.14／Double 100000／Boolean True／Date 同值／Double 12.5／Double 1.5 | 新增 |
| TC-CL-06 | 寫入 Empty 清除 | FR-CL-01 | P2 | A1="x"；`SetCellValue "Sheet1","A1",Empty`；讀 A1 | Empty | 新增 |
| TC-CL-07 | ClearCell 內容與格式 **ISS-10** | FR-CL-05 | P1 | 以 Excel 預先做 C3 粗體紅底並填值的範本；Open；ClearCell C3；存檔；以 Excel 檢視 | 內容清空（P1 通過條件）。現況預期：格式也被清除（與手冊不符）。期望行為：依人類決定（保留格式或修正手冊） | `13_Clear_Cell.vbs`（僅驗內容） |
| TC-CL-08 | VBS Integer 型別 **ISS-02** | FR-CL-06 | P1 | `SetCellValue "Sheet1","A1",123`（VT_I2）；`SetCellValue "Sheet1","A2",CByte(5)`；存檔；op2 讀回取 `TypeName`；另於 A3 寫公式需求以 Excel 檢查 A1 是否文字 | 現況預期：`TypeName = "String"`，值 "123"、"5"。期望行為：`TypeName = "Double"` | 新增 |
| TC-CL-09 | 非法位址與越界 | FR-CL-01、02 | P2 | `SetCellValue "Sheet1","1A","x"`；`SetCellValueByRC "Sheet1",0,1,"x"`；`…,1048577,1,"x"`；`…,1,16385,"x"` | 皆錯誤 | 新增 |
| TC-CL-10 | 公式與錯誤值 | FR-CL-03 | P3 | 以 Excel 預備 A1=1、A2=`=A1*2`、A3=`=1/0` 的檔案；Open 後讀 A2、A3 | A2=2（Double）；A3 為錯誤代碼字串 | 新增 |
| TC-CL-11 | Unicode 字串 | FR-CL-01 | P2 | 寫入「中文測試／日本語／emoji 😀」；存檔讀回 | 完全相同 | 新增 |
| TC-CL-12 | 讀空白儲存格 | FR-CL-03 | P2 | Create；讀 Z99 | `IsEmpty(v) = True` | 新增 |

## 6. 列與欄（TC-RC）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-RC-01 | 插入列 | FR-RC-01 | P1 | R1C1="Top"；`InsertRow "Sheet1",1`；讀 R1C1、R2C1 | R1C1 Empty；R2C1="Top" | `14_Insert_Row.vbs`（補 R1 為空檢查） |
| TC-RC-02 | 刪除列 | FR-RC-02 | P1 | R1="Keep"、R2="DeleteMe"、R3="Next"；DeleteRow 2；讀 R2C1 | R2C1="Next"；R1C1="Keep" | `15_Delete_Row.vbs`（補上移驗證） |
| TC-RC-03 | 列高 | FR-RC-03 | P1 | SetRowHeight 3, 25.5；Get；存檔後 op2 Get | 皆 25.5（±0.5） | `16_Set_Get_Row_Height.vbs`（補存檔後驗證） |
| TC-RC-04 | 插入欄 | FR-RC-04 | P1 | R1C1="ColA"；InsertColumn 1；讀 R1C2 | "ColA" | `17_Insert_Delete_Column.vbs` |
| TC-RC-05 | 刪除欄 | FR-RC-05 | P1 | 接 TC-RC-04；DeleteColumn 1；讀 R1C1 | "ColA" | `17_Insert_Delete_Column.vbs`（調整 Err 檢查順序，T-08） |
| TC-RC-06 | 欄寬 | FR-RC-06 | P1 | SetColumnWidth 2, 22.5；Get；存檔後 op2 Get | 皆 22.5（±0.5） | `18_Set_Get_Column_Width.vbs`（補存檔後驗證） |
| TC-RC-07 | 越界索引 | FR-RC-01～06 | P2 | 分別以 0、-1、1048577（列）、16385（欄）呼叫各方法 | 皆錯誤；活頁簿內容不變 | 新增 |
| TC-RC-08 | 公式隨插入調整 | FR-RC-01 | P3 | A1=1、A2=2、A3 預先公式 `=SUM(A1:A2)`（範本）；InsertRow 2；存檔以 Excel 檢視 | 公式自動調整為 `=SUM(A1:A3)` | 新增（手動） |

## 7. PDF 與列印（TC-OUT）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-OUT-01 | 匯出 PDF | FR-OUT-01 | P1 | ENV-A；Create＋A1；SaveWorkbook WB；`ExportToPdf WB, PDF` | PDF 存在、大小 > 0、可開啟且含 A1 內容；日誌有「開始 PDF 匯出」「成功」 | `09_Export_To_PDF.vbs` |
| TC-OUT-02 | 找不到 LibreOffice | FR-OUT-01、03 | P1 | ENV-C；`SetLibreOfficePath ""`；ExportToPdf | 錯誤，訊息含「PDF 匯出失敗」與「LibreOffice 路徑未設定或不存在」；日誌 ERROR | `09`（SKIP 分支） |
| TC-OUT-03 | 來源不存在 **ISS-14** | FR-OUT-01 | P2 | 新物件（無活頁簿）`ExportToPdf "NoSuch.xlsx", PDF` | 錯誤，訊息「PDF 匯出失敗：找不到 Excel 檔案」；現況 Err.Number 為通用值（記錄實際數值） | 新增 |
| TC-OUT-04 | 輸出更名與根目錄 | FR-OUT-01 | P2 | (a) 來源 `Report.xlsx` 輸出 `Other.pdf`；(b) 輸出到磁碟根目錄（如 `D:\TC_OUT_04.pdf`，測試後刪除） | (a) 只存在 `Other.pdf`，無殘留 `Report.pdf`；(b) 成功產出 | 新增 |
| TC-OUT-05 | 匯出覆寫來源 **ISS-06** | FR-OUT-04 | P1 | 準備 `A.xlsx`(A1="A") 與 `B.xlsx`(A1="B")；`OpenWorkbook A`；`ExportToPdf B, B.pdf`；op2 開 B 讀 A1 | 現況預期：B.xlsx 的 A1 變成 "A"，PDF 內容為 A。期望行為：B.xlsx 不變，PDF 內容為 B | 新增 |
| TC-OUT-06 | 覆寫既有 PDF | FR-OUT-01、NFR-08 | P2 | 目的 PDF 預先存在（記錄時間戳）後匯出 | 目的 PDF 被新檔取代 | 新增 |
| TC-OUT-07 | 逾時處理 | FR-OUT-05、NFR-05 | P2 | 以模擬程式（延遲 60 秒再結束）作為 `SetLibreOfficePath`；ExportToPdf 計時 | 約 30 秒後錯誤，訊息含「超時（30秒）」；模擬行程被終止 | 新增（需模擬程式） |
| TC-OUT-08 | 結束碼 0 但無輸出 **ISS-09** | FR-OUT-01 | P2 | 模擬程式立即以 0 結束、不產生檔案；ExportToPdf | 現況預期：無錯誤，PDF 不存在，日誌寫「成功」。期望行為：錯誤「未產生 PDF」 | 新增（需模擬程式） |
| TC-OUT-09 | 中文與空白路徑 | FR-OUT-01 | P2 | 來源與輸出位於 `TestData\測試 資料夾\` | PDF 正確產出 | 新增 |
| TC-OUT-10 | LibreOffice 已開啟 **ISS-15** | FR-OUT-01 | P3 | 先手動開啟 LibreOffice Calc（任意檔）後執行 TC-OUT-01 | 記錄實際結果（成功／無輸出／卡住） | 新增（手動） |
| TC-OUT-11 | 預設印表機列印 | FR-OUT-02 | P1 | 預設印表機設為 Microsoft Print to PDF；`PrintWorkbook WB, ""` | 無錯誤；日誌「列印命令完成」；虛擬印表機產生輸出（可能跳出存檔視窗，需人工確認） | `10_Print_WorkBook.vbs`（修正 T-02） |
| TC-OUT-12 | 指定印表機 | FR-OUT-02 | P2 | `PrintWorkbook WB, "Microsoft Print to PDF"` | 同 TC-OUT-11 | 新增 |
| TC-OUT-13 | 不存在的印表機 | FR-OUT-02 | P3 | `PrintWorkbook WB, "NoSuchPrinter"` | 記錄實際結果（LibreOffice 可能結束碼 0 而不報錯） | 新增 |
| TC-OUT-14 | 自動偵測路徑順序 | FR-OUT-03 | P2 | (a) 僅 PATH 有 Portable soffice；(b) 僅 Program Files 安裝版；(c) Set 覆寫指向不存在檔 + 安裝版存在 | 日誌與結果顯示：(a) 使用 PATH；(b) 使用 Program Files；(c) 回退到安裝版且成功 | 新增 |
| TC-OUT-15 | 列印前自動存檔 | FR-OUT-04 | P2 | Open WB；改 A1；`PrintWorkbook WB, ""`；op2 讀 WB | WB 的 A1 已更新（現況行為） | 新增 |

## 8. 設定與日誌（TC-CFG）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-CFG-01 | 設定 LibreOffice 路徑 | FR-CFG-01 | P1 | ENV-D；SetLibreOfficePath Portable 路徑；執行 TC-OUT-01 | 成功 | `20_Config_Functions.vbs`＋`09` |
| TC-CFG-02 | 日誌資料夾與檔名 | FR-CFG-02、05 | P1 | SetLogDirectory `TestData\Logs\TC02`（不存在）；SetLogEnabled True；SetLogLevel 3；ExportToPdf | 資料夾自動建立；檔名 `ExcelKit1.2.4.0_<yyyyMMdd>.txt`；行格式 `[yyyy-MM-dd HH:mm:ss] [INFO] …` | `20`（補檔案驗證，T-07） |
| TC-CFG-03 | 停用日誌 | FR-CFG-03 | P2 | 新資料夾；SetLogEnabled False；ExportToPdf | 不產生日誌檔 | 新增 |
| TC-CFG-04 | 等級過濾 | FR-CFG-04、05 | P2 | SetLogLevel 1；執行成功匯出與一次失敗匯出（TC-OUT-03） | 僅有 `[ERROR]` 行 | 新增 |
| TC-CFG-05 | 預設日誌路徑無權限 **ISS-18** | FR-CFG-02 | P3 | 將 DLL 放在唯讀資料夾註冊，未呼叫 SetLogDirectory；ExportToPdf | 匯出成功不受影響；現況預期：無日誌且無提示 | 新增 |
| TC-CFG-06 | 跨物件共用設定 **ISS-13** | 2.4-2 | P2 | op 設 SetLogEnabled False；op2（同一行程）執行 ExportToPdf | 現況預期：op2 也不寫日誌。期望行為：依人類決策（行程層級或物件層級），並寫入手冊 | 新增 |
| TC-CFG-07 | 等級夾值 | FR-CFG-04 | P3 | SetLogLevel 0 後觸發 INFO 事件；SetLogLevel 9 後觸發 INFO | 0 → 僅 ERROR；9 → 有 INFO | 新增 |
| TC-CFG-08 | 清除路徑覆寫 | FR-CFG-01 | P3 | SetLibreOfficePath 有效路徑後再 SetLibreOfficePath ""；ENV-A 匯出 | 回到自動偵測並成功 | 新增 |

## 9. 整合（TC-INT）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 預期結果 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-INT-01 | LabVIEW 讀取既有檔 | 2.2 | P1 | ENV-E 執行 `Reference/LabVIEW/ExcelKit_ReadExistFile.vi` | VI 無錯誤並顯示正確資料 | 既有 VI |
| TC-INT-02 | LabVIEW PDF＋列印 | 2.2 | P1 | ENV-E 執行 `ExcelKit_SavePDFPrint.vi` | PDF 產出、列印命令完成 | 既有 VI |
| TC-INT-03 | PowerShell 驗證 | 手冊 3 章 | P3 | `New-Object -ComObject ExcelKit.ExcelOperator`；`GetVersionInfo()` | 回傳版本字串 | 新增 |
| TC-INT-04 | 範例腳本端對端 | 手冊 7 章 | P2 | 修改 `Reference/VBS/example.vbs` 路徑至測試資料夾後執行 | 存檔與 PDF 成功 | `Reference/VBS/example.vbs`、`Test_PDF.vbs` |

## 10. 效能與資源（TC-PERF，僅量測記錄）

| 編號 | 名稱 | 需求 | 優先 | 步驟 | 記錄項目 | 既有腳本 |
| --- | --- | --- | --- | --- | --- | --- |
| TC-PERF-01 | 大量寫入 | NFR-04 | P3 | 以 SetCellValueByRC 寫入 1,000 列 × 10 欄後存檔 | 寫入與存檔耗時、檔案大小 | 新增 |
| TC-PERF-02 | 重複建立 **ISS-04** | NFR-04 | P3 | 同一物件連續 Open 同一檔 200 次（不釋放） | 行程記憶體前後變化 | 新增 |

---

## 附錄 A：既有腳本對應總覽

| 腳本 | 主要案例 | 需補強 |
| --- | --- | --- |
| 01_Reference_COM_Component | TC-ENV-01 | — |
| 02_Read_DLL_Version | TC-VER-01 | 版本字串比對、DLL 名稱 |
| 03_Create_WorkBook | TC-WB-01 | 工作表名稱驗證 |
| 04_Create_WorkSheet | TC-SH-01 | 名稱驗證 |
| 05_Set_Cell_By_Name | TC-CL-01 | 移除 msgbox；不依賴 ACE |
| 06_Set_Cell_By_Row_Column | TC-CL-02 | 不依賴 ACE |
| 07_Read_Cell_Value | TC-CL-03 | 型別檢查 |
| 08_Save_WorkBook | TC-WB-02 | — |
| 09_Export_To_PDF | TC-OUT-01、02 | 檢查 PDF 大小 |
| 10_Print_WorkBook | TC-OUT-11 | 移除 msgbox 與硬編碼印表機 |
| 11_Open_WorkBook | TC-WB-03 | — |
| 12_Delete_Sheet | TC-SH-04 | 驗證已刪除 |
| 13_Clear_Cell | TC-CL-07 | 格式檢查 |
| 14_Insert_Row | TC-RC-01 | R1 空白檢查 |
| 15_Delete_Row | TC-RC-02 | 上移驗證 |
| 16_Set_Get_Row_Height | TC-RC-03 | 存檔後驗證 |
| 17_Insert_Delete_Column | TC-RC-04、05 | Err 檢查順序 |
| 18_Set_Get_Column_Width | TC-RC-06 | 存檔後驗證 |
| 19_Get_All_Sheet_Names | TC-SH-09 | 內容驗證 |
| 20_Config_Functions | TC-CFG-01、02 | 日誌檔驗證 |

## 附錄 B：案例統計
| 群組 | 案例數 | P1 | P2 | P3 |
| --- | --- | --- | --- | --- |
| BLD | 3 | 1 | 2 | 0 |
| ENV | 4 | 2 | 1 | 1 |
| VER | 2 | 1 | 1 | 0 |
| WB | 11 | 5 | 6 | 0 |
| SH | 9 | 3 | 5 | 1 |
| CL | 12 | 7 | 4 | 1 |
| RC | 8 | 6 | 1 | 1 |
| OUT | 15 | 4 | 9 | 2 |
| CFG | 8 | 2 | 3 | 3 |
| INT | 4 | 2 | 1 | 1 |
| PERF | 2 | 0 | 0 | 2 |
| **合計** | **78** | **33** | **33** | **12** |

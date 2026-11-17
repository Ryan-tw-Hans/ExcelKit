# 軟體測試計畫（STP）｜ExcelKit v1.2.4

| 項目 | 內容 |
| --- | --- |
| 文件編號 | STP-ExcelKit-v1.2.4 |
| 文件版本 | v0.1（草稿，待人類確認） |
| 日期 | 2026-11-14（Asia/Taipei） |
| 撰寫者 | Claude（測試工程師角色；自查，非獨立審查） |
| 受測版本 | ExcelKit 1.2.4（`ExcelKit v1.2.4 (2026-04-05)`） |
| 依據 | `Requirements/SRS-ExcelKit-v1.2.4.md`、`Requirements/CodeReview-ExcelKit-v1.2.4.md` |
| 配套文件 | `TestPlans/TCS-ExcelKit-v1.2.4.md`（測試案例）、`TestPlans/RTM-ExcelKit-v1.2.4.md`（追溯矩陣） |

## 1. 目的
建立 ExcelKit 1.2.4 的完整測試基線：驗證 SRS 所列 COM 功能、確認程式碼審查發現的風險（ISS-xx）、補足現有 20 支 VBS 腳本的覆蓋缺口，並作為後續修正版本的回歸依據。

## 2. 測試範圍

### 2.1 納入
| 範圍 | 需求群組 |
| --- | --- |
| 建置與部署（net472、net10.0、Costura、regasm） | NFR-01～03 |
| COM 註冊與物件建立（x64、x86） | NFR-01、4.1 |
| 版本資訊 | FR-VER-01、NFR-09 |
| 活頁簿 建立／開啟／儲存 | FR-WB-01～03 |
| 工作表 新增／刪除／列舉／自動建立 | FR-SH-01～04 |
| 儲存格 讀寫／清除／型別對應 | FR-CL-01～06 |
| 列與欄操作 | FR-RC-01～06 |
| PDF 匯出、列印、LibreOffice 路徑、逾時、自動存檔 | FR-OUT-01～05 |
| 設定與日誌 | FR-CFG-01～05 |
| 整合（LabVIEW 範例 VI、PowerShell、example.vbs） | 2.2 使用情境 |
| 效能基準與資源（量測記錄，無門檻） | NFR-04 |

### 2.2 不納入
- .NET 專用 API（`FormatCell`、三參數 `ExportToPdf`、`Dispose`）之功能測試：COM 用戶無法呼叫，僅以程式審查覆蓋（FR-CL-07、FR-OUT-06、FR-RES-01）。
- LibreOffice 轉換的版面正確性（字型、分頁），僅確認檔案產出與可開啟。
- 實體列印品質；以「Microsoft Print to PDF」虛擬印表機代替實體印表機確認列印路徑。
- 多執行緒壓力測試（NFR-07 未設計執行緒安全，僅文件約束）。

## 3. 測試策略

| 層級 | 方法 | 工具 | 對應 TCS 群組 |
| --- | --- | --- | --- |
| 建置驗證 | `dotnet build -c Release`，檢查兩個目標框架輸出 | .NET SDK／VS | TC-BLD |
| 環境／部署 | regasm 註冊、登錄檢查、x86／x64 建立物件 | regasm、PowerShell、`%SystemRoot%\SysWOW64\cscript.exe` | TC-ENV |
| COM 功能（正向） | VBS 腳本逐 API 驗證，以「寫入→存檔→以另一物件開啟讀回」確認持久化 | cscript、既有 `測試腳本/` | TC-WB／SH／CL／RC |
| 負面與邊界 | 不存在檔案、非法名稱、越界索引、未初始化 | 新增 VBS | 各群組負面案例 |
| 外部程式 | LibreOffice 正常、缺少、逾時、無產出；以模擬 soffice 程式製造異常 | LibreOffice Portable、模擬程式 | TC-OUT |
| 設定與日誌 | 檢查日誌檔名、內容、等級過濾 | VBS＋檔案檢查 | TC-CFG |
| 整合 | 執行 `Reference/LabVIEW/*.vi`、`Reference/VBS/example.vbs` | LabVIEW、cscript | TC-INT |
| 效能基準 | 計時與記憶體量測，僅記錄 | PowerShell `Measure-Command`、工作管理員 | TC-PERF |
| 回歸 | 修正任何 ISS 後重跑所有 P1＋受影響群組 | `Run_All_Ordered_Tests.vbs`（整理後） | 全部 P1 |

### 3.1 已知問題案例的判定方式
TCS 中標示「ISS-xx」的案例同時列出「現況預期」與「期望行為」：
- 對 1.2.4 執行時，結果符合「現況預期」→ 記為 **失敗（已知 ISS-xx）**，用以確認問題存在；不符合現況預期 → 記為「失敗（新發現）」並回報 PM。
- 修正版本執行時，以「期望行為」判定通過與否。期望行為須先經人類核准（見 CR 第 4 節）。

### 3.2 驗證原則
- 寫入類 API 一律以**另一個 ExcelOperator 物件重新開啟**讀回驗證，不只看記憶體。
- 不依賴 ACE OLEDB 作為唯一驗證手段（CR T-09）；可另以 Excel／LibreOffice 人工開啟作為輔助證據。
- 每個案例使用獨立的測試檔名或在開始時清除，避免案例間相依。

## 4. 測試環境

| 編號 | 作業系統 | 呼叫端位元 | .NET | LibreOffice | 用途 |
| --- | --- | --- | --- | --- | --- |
| ENV-A | Windows 10/11 x64 | x64（`System32\cscript.exe`） | 4.7.2+ | 安裝版（Program Files） | 主要功能測試 |
| ENV-B | 同上 | x86（`SysWOW64\cscript.exe`） | 4.7.2+ | 安裝版 | 位元相容 |
| ENV-C | 同上 | x64 | 4.7.2+ | 未安裝、PATH 無 soffice | LibreOffice 缺少情境 |
| ENV-D | 同上 | x64 | 4.7.2+ | Portable 版＋`SetLibreOfficePath` | 免安裝情境 |
| ENV-E | 有 LabVIEW 的量測站 | 依 LabVIEW 位元 | 4.7.2+ | 依現場 | 整合測試 |

- 測試資料：`測試腳本/TestData/`；日誌輸出：`測試腳本/TestData/Logs/`。
- 每次測試記錄：OS 版本、LibreOffice 版本、ExcelKit DLL 檔案版本、註冊位元。
- 實際作業系統與 LibreOffice 版本尚未確認，執行前由測試人員填入。

## 5. 進入與離開條件

### 5.1 進入條件
1. `dotnet build` 兩目標無錯誤（TC-BLD-01 通過）。
2. 受測 DLL 已以對應位元 regasm 註冊，`GetVersionInfo` 回傳 `ExcelKit v1.2.4 (2026-04-05)`。
3. 既有 VBS 腳本已移除互動 `msgbox`、修正預設 DLL 名稱（CR T-01～T-03），或執行者確認以手動方式處理。

### 5.2 離開（通過）條件
1. 所有 **P1** 案例通過；已知 ISS 案例皆為「失敗（已知）」或已修正通過，無「失敗（新發現）」未處理。
2. P2 案例通過率記錄完成；未通過者皆已登錄 IssueLog 並有處置決定。
3. 「未執行」案例需列出原因（例如無 LabVIEW 環境），**不得計為通過**。
4. 測試報告完成並經 PM 確認。

### 5.3 暫停與恢復
- 暫停：COM 無法建立（TC-ENV-01 失敗）、建置失敗、測試環境無法提供 LibreOffice 以外的必要元件。
- 恢復：阻塞原因排除後，從 TC-ENV-01 重新開始。

## 6. 優先度定義
| 等級 | 定義 |
| --- | --- |
| P1 | 核心流程與資料正確性；任一失敗即不可發布 |
| P2 | 錯誤處理、邊界、設定；失敗需評估 |
| P3 | 效能基準、文件一致性、輔助驗證 |

## 7. 風險與因應

| 風險 | 影響 | 因應 |
| --- | --- | --- |
| LibreOffice 版本差異導致輸出不同 | PDF 案例結果不穩 | 固定測試用 Portable 版本並記錄版本號 |
| 印表機環境不一致 | 列印案例無法重現 | 使用 Microsoft Print to PDF；實體列印列為人工驗收 |
| ACE OLEDB 未安裝或位元不符 | 05／06 誤判 | 改以 ExcelOperator 讀回為主要驗證 |
| GUID 佔位值與他元件衝突（ISS-07） | 建立到錯誤元件 | TC-ENV-03 先檢查登錄 |
| 測試會覆寫檔案（ISS-06） | 誤刪測試以外的檔案 | 一律在 `TestData` 下操作，禁止指向實際資料 |
| LabVIEW 環境不可得 | 整合案例未執行 | 標記「未執行」並註明，由現場人員補測 |

## 8. 測試結果記錄格式
依專案 test-plan 規範，每個案例記錄：

| 欄位 | 說明 |
| --- | --- |
| 案例編號 | TC-xxx-nn |
| 結果 | 通過／失敗（已知 ISS-xx）／失敗（新發現）／未執行／不適用 |
| 執行方法 | 腳本名稱或手動步驟 |
| 環境 | ENV-A～E |
| 受測版本 | DLL 檔案版本＋`GetVersionInfo` 字串 |
| 證據位置 | 報告、日誌、輸出檔路徑 |
| 執行者／日期 | — |

報告檔：`TestReports/TR-ExcelKit-v1.2.4-<日期>.md`；自動化腳本另產出 `測試腳本/TestReports/TestReport_COM_<版本>.md`。

## 9. 交付項目
1. 本測試計畫。
2. 測試案例規格（TCS）。
3. 需求追溯矩陣（RTM）。
4. 新增／整理後的 VBS 測試腳本（待分派）。
5. 測試報告（執行後產出）。

## 10. 目前狀態
本計畫為草稿，**所有案例均未執行**。腳本整理與新增負面案例需經 PM 受理並分派後進行。

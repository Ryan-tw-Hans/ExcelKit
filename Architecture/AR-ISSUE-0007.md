# ISSUE-0007 打包版本與內容架構評估草案 v0.1

SA／系統整合／維運：Codex 依序執行，自查非獨立審查。依據目前工作區 Installer/Package-ExcelKit.ps1、ExcelKit.iss、Startup.iss、Verify-ExcelKit.ps1、Installer/FileVersion.ps1 與 publish/Installer/Setup 的既有報告；包含既存未提交工作。需求：Requirements/RS-ISSUE-0007.md。狀態：待核准。

## 現況
打包器以 `1.2.4+yyyyMMddHHmmss` 作安裝器版本，檔名改為 `ExcelKit-Setup-1.2.4-yyyyMMddHHmmss-TEST.exe`，以及對應 manifest/build JSON；完整 Reference/Media 被複製並加入清單。套件版本會寫入 Inno AppVersion、安裝選擇記錄與 manifest；Verify-ExcelKit 比較選擇紀錄、已安裝版本與 manifest。DLL 版本與安裝器版本獨立，DLL 檔案版本台帳另行管控。

## 方案及取捨
| 方案 | 優點 | 風險 |
| --- | --- | --- |
| A：安裝器版本與檔名統一使用 `<基準>-yyyyMMdd-NN`，排除單一影片 | 符合使用者在畫面及檔名看到的版本，清單與安裝記錄一致 | 必須兼容舊版字串、同日序號配號及並行控制；舊版用戶端仍需測試 |
| B：只改檔名，內部仍使用秒級時間戳 | 改動較少，沿用原安裝版本契約 | 介面與報告仍顯示秒級版本，不完全符合打包版號要求 |

推薦 A。DLL 仍用既有秒級產品版本與遞增 FileVersion，避免混淆 COM 與正式版；並不因移除 TEST 而轉成正式版。素材由源頭過濾，禁止把排除影片送進 staging 或清單。保留原影片與舊產物，不以安裝升級刪除既有影片。只移除使用者指定的檔案。

## 預防衝突與相容
同日序號應基於可驗證的既有輸出與保留紀錄，以同一台工作區互斥鎖避免併發；EXE、manifest、build 只要一項已存在即不得覆寫。測試失敗保留記錄與中間成果、不得回收曾分配序號。多機／多工作區不可能單靠本地掃描保證唯一，需集中協調或安全停止。預檢所顯示的版本須能在正式打包時保持相同且未被其他程序搶用；不可靜默換號。

新安裝包須能識別舊版以時間戳命名的已安裝版本，不重寫舊歷史；檢查中英版本顯示、安裝選項紀錄、驗證腳本及升級／回復。安裝內容以清單檢驗；舊版已安裝影片可能仍留存，不承諾清除。舊檔案與原始影片留存可作回復來源，但實際回復需 Sandbox 驗證。無新增網路、權限或常駐服務；交付位置不變，發行方式與版本契約變更。

## 測試及決策
最少覆蓋同日連續、跨日、重複／部分產物、預檢至打包間衝突與並行、舊版時間戳升級與回復、x86/x64 COM、排除影片及其他素材、清單/哈希、中英顯示與安裝後驗證。Windows Sandbox 允許的情況下實測，無法執行不得算通過。可採用條件：PM 記錄人類對需求、此架構方案 v0.1 的核准並完成測試計畫；實作後 SA 自查待執行。UI 只改版本文字值，不改版面；實際顯示需測試。DB schema 不適用；維運需記錄配號、異常處理與回復，責任由 PM 與維運角色追蹤。

## 2026-11-17 使用者追加：專案 skill 同步（草案 v0.2）
使用者要求將規則同步至腳本與 skills。由於既有 `.agents/skills/excelkit-package/SKILL.md` 明寫安裝器秒級時間戳及 `-TEST.exe`，`.agents/skills/build-release/SKILL.md` 與 `AgentUse/skills/build-release/SKILL.md` 提供一般時間戳規則，應針對 ExcelKit 安裝器加入日期序號的專案特例；DLL 仍用已核准規則。同步 Installer/README 與測試，不改未受指示的全域 skill 或根目錄 agent-workflow 發行包。使用者明確指定的本專案 Agent 控制資料可修改，不因此授權提交；實際寫入仍受環境權限保護。評估、測試及核准均以 v0.2 為準。

## 2026-11-17 決策：檔名使用 DLL FileVersion（v0.3）
使用者明確選用 `ExcelKit-Setup-<DLL FileVersion>-yyyyMMdd-NN.exe`，取代 v0.1 的三段基準名稱；安裝器 AppVersion／manifest／build report 的 installer_version 採相同四段 DLL FileVersion + 日期序號。軟體須以實際打包 DLL FileVersion 檢核名稱，維持舊時間戳紀錄可讀；不同 FileVersion 但相同日期的包仍以同一日期序號序列防止重用。同工作區互斥及已保留配號不可回收，多工作區需集中協調。產品名稱不含 TEST 並不構成正式發布；發行類型仍為測試。使用者以具體指令核准此方案並要求腳本與 skills 同步，不等於核准合併或對外發布。

## 實作後自查（2026-11-17）
Codex 自查：採用使用者確認的 `<DLL FileVersion>-yyyyMMdd-NN`，一工作區內以互斥與保留紀錄防止重複，歷史舊格式不改。來源排除單一影片，其他素材保留；新包 manifest、版本、SHA-256、同日序號、負面驗證及 Sandbox 舊版升級／x86/x64 COM／回復通過，細節見 TestReports/TR-ISSUE-0007.md。跨工作區及機器需自行協調，不聲稱全域唯一；本版 Sandbox 中英文歡迎頁文字檢查通過；既有影片保留情境未實測。無新的網路、權限或正式版發布。

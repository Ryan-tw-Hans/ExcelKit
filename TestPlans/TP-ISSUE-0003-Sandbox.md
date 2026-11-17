# Windows Sandbox 驗證準備｜ISSUE-0003

執行者：Codex（PM／SA／測試，自查）；2026-11-15。使用者表示可使用 Windows Sandbox，不需自行建立虛擬機。

## 本次準備與已完成檢查
- Windows 11 專業版；Windows Sandbox 原為停用。已用管理員權限啟用 `Containers-DisposableClientVM`，使用 `-NoRestart`；系統回傳 RestartNeeded=True，尚未啟動 Sandbox。
- 將舊 DLL 複製至專案隔離測試輸入資料夾，與原檔 SHA-256 均為 `B2C2FC06817EA615CED43C90D910F259E516E4A10BEE3941DC701EB1339FA84E`。原檔及主機 COM 註冊未修改。
- 32／64 位元 RegAsm 使用 `/regfile /codebase` 產生唯讀分析結果；兩檔雜湊相同。它不涵蓋 TypeLib、Interface 及自訂註冊函式，不能作完整備份。主機的既有 TypeLib 登錄指向兩個目前不存在的暫存檔案，回復設計須保留這個既存狀態及限制。
- 已準備 `.tools/installer-tests/sandbox/ExcelKit-LegacyBaseline.wsb`：網路停用、輸入唯讀，只有專用結果資料夾可寫；Sandbox 內以舊 DLL 副本建立測試註冊、移除、重新註冊，記錄指定 COM 登錄樹並比對測試狀態。腳本限制 WDAGUtilityAccount 執行。PowerShell 語法及 WSB XML 已靜態檢查；尚未在 Sandbox 實際執行。

## 後續與驗收限制
重新開機後確認 Sandbox 可啟動，再執行基礎測試並核對結果。這只是指定樹的基礎測試，無法自行證明完整登錄變更範圍或產品安裝器的失敗回復；還須設計並測試全範圍快照、失敗注入、中斷、x86／x64 COM 客戶端及實際安裝／移除。依已核准的舊版處理方案，完成這些驗證前不可宣稱自動移除安全或完成。

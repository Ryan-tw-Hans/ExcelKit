# 可重複打包與安裝驗證測試結果｜ISSUE-0003

受測來源為 feature/ISSUE-0003-installer 工作區；Codex 依序扮演軟體、測試與 SA，為自查，非獨立審查。測試計畫 TP-ISSUE-0003-repeatable-package.md。

- 通過：Windows PowerShell 5.1 執行預檢及完整乾淨快照打包；.NET SDK 10.0.403，Inno Setup 7.1.1，NuGet 鎖定還原。正式基準 1.2.4；測試安裝器版本 `1.2.4+20261116141740`，DLL 產品版本 `1.2.4+20261115094510`，Assembly／FileVersion `1.2.4.0`。最終安裝檔 SHA-256 `5215D0036D7D17AF4B45A6E5F6B2BB01632C8861C044FDFB2A2E76FF830D5015`；安裝清單 SHA-256 `F4187688F11BC203FA802B59A3183D6AC48949A3F44D8C04DEB9260E57631C00`。
- 通過：獨立核對建置報告中的 27 個來源及清單內 25 個打包檔案 SHA-256；exe、清單、預處理腳本與中繼資料雜湊一致。打包時從兩框架實際 DLL 核對 Assembly、FileVersion 與 ProductVersion；安裝器預處理中繼資料包括 DLL 產品及組件版本。
- 通過：隔離副本預檢不產生輸出；不存在編譯器、來源版本不符、必要素材缺失均停止；同一版本重試被拒，既有 exe 雜湊不變。測試證據 .tools/package-preflight-tests/7fd0b8b04244483ab37ed46c03fa75c7。
- 通過（隔離檔案樣本，非真實安裝）：必要檔案雜湊、修改檔案回報失敗、失敗時不執行 Functional、缺少已選檔案回報失敗、缺選項紀錄回報 unknown、路徑越界拒絕。證據 .tools/package-tests/f745f6e76bb24f75be1c902823e80a0d。
- 通過：skill-creator quick_validate（使用專案 .tools 隔離 PyYAML 與 UTF-8 模式）及 UI metadata UTF-8 檢查。既有 Rebuild-Setup.ps1 已驗證預檢轉呼叫標準入口。
- 過程修正：首次打包因 Windows 換行判斷而安全停止；其後 Windows PowerShell 5.1 將 ISCC stderr 的既有警告視為錯誤而停止。保留的 staging 未動既有產物；修正判斷與程序輸出處理後最終流程通過。SDK 建置有 23 項既有程式警告、0 錯誤；Inno 有既有 FAULT 未使用提示。
- 未執行：使用者互動及靜默安裝、安裝選項狀態檔的實際產生、x86/x64 COM 功能、升級／失敗回滾與解除安裝。Windows Sandbox 本次已啟動但長時間未產生結果檔，已確認並關閉本次工作階段；未推論阻擋原因。先前版本曾被 Application Control 阻擋，不代表本次實測結果。
- 應用 DB 無影響；UI 未修改現有歡迎頁設計，但安裝器新增安裝狀態記錄，仍需測試。SA／系統整合／維運自查有條件採用，真實安裝與回復未驗證，整體驗收、維運就緒及合併前置條件未完成。
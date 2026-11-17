# 啟動頁架構評估｜ISSUE-0003

Codex 擔任 SA、系統整合及維運，自查；2026-11-16。已由使用者以「開工」確認 Requirements/RS-ISSUE-0003-startup-links-draft.md，保留 L3。輸入為 Installer/ExcelKit.iss、ComTransaction.ps1 與 publish/bin/net472/ExcelKit.dll；資料庫不適用。

現況是 InitializeSetup 在精靈前檢查 Framework；PrepareToInstall 管控 COM Inspect、遷移同意與交易。提案在歡迎頁唯讀查詢兩種位元視圖的 COM 登錄及安裝器登錄，連結由原使用者身分開啟；移至 PrepareToInstall 最前方檢查 Framework，先於 ExtractTemporaryFile、Inspect、Begin。原交易函式不依版本顯示結果決策。

維持現況無法提供啟動入口；另建外部工具會新增交付及維運邊界。推薦原生 Inno 控制項及唯讀登錄查詢。打包 DLL 的組件版本由 build 前從實際 DLL 取得，編譯時用 SHA-256 比對，防止過期中繼資料；無法確認檔案、版本、權限或安裝目錄時顯示未知，不據此執行安裝決策。註冊組件版本與實際檔案版本分開呈現，只有雜湊相等才聲稱檔案一致。未註冊不等同系統沒有 DLL。

系統整合與維運條件：不得在資訊頁執行 DLL、寫入 COM 或自動下載；無效／衝突的安裝路徑不得開啟；保留安裝前交易備份和失敗回復。維持原 UAC 權限提示，外部瀏覽器或檔案總管以原使用者權限開啟。Framework 缺失與靜默安裝阻擋、權限拒絕、多筆登錄、x86/x64、UI 縮放及原有 COM 回歸需要隔離 Windows 實測。僅編譯、靜態檢查不能證明維運就緒。

結論：有條件採用，使用者已核准範圍；必須通過 TestPlans/TP-ISSUE-0003-startup-links.md。實作後符合性與維運就緒：待實測。固定輸出 publish/Installer/Setup；正式基準 1.2.4 不變。
## 實作後自查

受評版本 1.2.4+20261116121333：來源及建置設定顯示已依提案加入歡迎頁、唯讀註冊資訊、DLL 雜湊與安裝前 Framework 檢查；編譯、靜態及中繼資料負向測試見 TestReports/TR-ISSUE-0003-startup-links.md。GUI、缺 Framework、不同登錄情況、安裝及回復均未實測；架構符合性與維運就緒整體未確認。

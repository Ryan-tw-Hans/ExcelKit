# ISSUE-0008｜GitHub Release 附加安裝檔需求草案 v0.2

狀態：使用者已核准 v0.2；PM：Codex。使用者更正：希望從 GitHub Release 取得 Setup 安裝檔。原 v0.1 只上傳 Actions artifact，並明列不建立 Release，未達此需求。本次為 L3 範圍變更，涉及公開發布介面與 GitHub 權限。現行 GitHub API 未列出 Release，既有序號 tag 的 Source code ZIP 是 GitHub 自動封存，而非 Setup。

## 建議範圍
1. 既有 Actions 測試、編譯、序號保留及 manifest 檢查成功後，新增 Release 發布步驟，使用該次已保留 tag，明確標示 prerelease／測試版，附加該次 EXE、manifest、build report。僅在檔案齊全、版本及雜湊一致時公開 Release；失敗保留待診斷狀態，不覆寫或刪除既有資產。
2. 將建置當次來源 SHA、安裝器版本及 SHA-256 寫入 Release 說明，提醒雲端建置不等於安裝、升級、COM、GUI、LabVIEW 驗收。Release 為公開儲存庫的公開測試版，並非正式版，不變更正式版基準。
3. 建議本次先公開目前已建置的 1.2.4.3-20261117-03 測試版為 prerelease，附加其三個產物；往後成功打包自動建立對應 prerelease。GitHub 仍會顯示其自動產生的 Source code ZIP/TAR，使用者應點選以 `.exe` 結尾的安裝檔。
4. 同步保留 Actions artifact，對重試及同名資產衝突採安全停止、人工診斷，不自動覆蓋。既有分配序號不得刪除或重用。使用者另行明確要求正式版前不建立正式 Release。

## 適用性
架構適用（發布方式與信任邊界）；UI 視覺不適用（無自建介面）；資料庫不適用；系統整合適用（Actions、Release API、產物）；維運適用（公開資產、重試、刪改與可追溯）。SA 評估 `Architecture/AR-ISSUE-0008-release-v02.md`，測試草案 `TestPlans/TP-ISSUE-0008-release-v02.md`。使用者已核准此版本方案，實作與驗證進行中。

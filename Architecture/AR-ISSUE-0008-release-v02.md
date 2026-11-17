# ISSUE-0008｜Release 附加檔架構評估草案 v0.2

SA／系統整合／維運：Codex 依序執行，自查非獨立審查。需求 `Requirements/RS-ISSUE-0008-release-v02.md`。現況：Windows Actions 建置並上傳 artifact；GitHub release 清單為空，保留序號 tag 會顯示 GitHub 原始碼封存。原 v0.1 排除 Release；本變更新增公開交付，需確認。

| 選項 | 效果與取捨 |
| --- | --- |
| A：通過測試後以現有 tag 建立 prerelease 並附加 EXE、manifest、build report | Releases 可直接下載 Setup；自動化可追溯，需嚴格處理公開權限、重試與半完成狀態 |
| B：維持 Actions artifact | 建置已可下載，但不滿足使用者在 Release 找 Setup 的需求 |
| C：正式 Release | 可直接下載，但目前安裝與 COM 尚未完整驗收，且使用者尚未要求正式版 |

建議 A。先建立 draft，完成上傳並以 GitHub API 核對資產名稱、大小及可用的 SHA-256 與預期一致，再轉為 prerelease；避免未上傳完整的公開 Release。需要測試 Actions contents:write、現有受保護 tag 與 Release API 的相容性；禁止改寫 tag。重新執行遇到同 tag、已存在資產或狀態不明時安全停止；預設不刪除、覆寫或重新發布，需提供可診斷的復原程序。不同 build 須使用已保留的不同序號。必須只允許受信任的 push／手動執行，禁止 untrusted pull request 使用寫入憑證。公開 prerelease 發生在成功打包後，檢查及 release 發布分開記錄；Actions artifact 過期不影響 release asset，但資產遭刪改時需可追溯與回報。GitHub 自動的 Source code 封存無法視為安裝包。

待確認既有 workflow 每次 feature push 均觸發，A 也會因此每次產生公開 prerelease；使用者已核准此行為。現有 03 的 artifact 可供第一次發布但需先取得三個原產物並核對雜湊。正式安裝、COM 等驗收仍未執行。

# ISSUE-0008｜Release 附加檔測試草案 v0.2

測試工程師：Codex，自查非獨立審查。待方案核准後定案；目前未執行。
- 測試完整成功路徑：依指定 commit 及唯一序號完成建置，draft 附上 EXE、manifest、build JSON；核對版本、檔名、長度與 SHA，轉為 prerelease 後從 GitHub Release 下載 EXE 並與 build report 比對。
- 模擬缺檔、雜湊不符、上傳中斷、API 權限不足、重複 release／asset、重跑及並發，確認不覆蓋既有資產、不改寫 tag，且不公開半成品；記錄狀態不明的恢復操作。
- 驗證不可信 PR 不具寫入能力；確認 Release 的 prerelease 標記、來源 commit、下載入口，Source code ZIP 不被誤認為安裝檔。
- 本次新增既有 03 測試版發布時先核對來源與現有 artifact；執行後核對實際下載內容與 SHA。安裝、COM、升級／回復驗收獨立，未執行不能算通過。

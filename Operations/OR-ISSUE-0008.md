# ISSUE-0008｜維運與導入檢查紀錄

2026-11-17：Codex 依序以 PM、SA、系統整合及維運角色執行，自查非獨立審查。需求與架構 v0.1 已由使用者核准。

## 導入前檢查
- `origin` 為 `Ryan-tw-Hans/ExcelKit`；查詢 `refs/tags/excelkit-package/*` 未發現現有保留參照。GitHub 儲存庫 rulesets 查詢結果為空，故目前未具備核准方案要求的刪改保護。GitHub 連接器目前無建立 ruleset 工具，Git Credential Manager 非互動查詢未完成並已中止；未變更遠端設定。
- 此工作區日期序號歷史：`.tools/package-versions/20261117-01.json` 及同版 EXE、manifest、build report。其他工作區或移出產物是否存在，待使用者確認；不能推定目前盤點涵蓋所有歷史。
- 來源檔均以 CRLF 存放，FileVersionLedger 的來源雜湊依實體位元組計算；Git 帳面正規化可能轉換行尾。雲端 checkout 必須驗證還原位元組與版本台帳相符，不可忽略雜湊防護。
- 既有本機 preflight 測試在有效案例失敗，原因為讀取主機已註冊的同版 DLL，但測試 fixture 沒有匹配 manifest。未變更主機登錄；必須在隔離或合理設計的 fixture 中驗證。

## 未完成與阻擋
- 需建立並核實針對 `refs/tags/excelkit-package/**/*` 的 active tag ruleset，限制更新與刪除且不設 bypass；核對有效目標及後續設定變更的偵測策略。具有管理權限者仍可能改動規則本身，維運需監控；不應聲稱防止所有管理員操作。
- 需確認跨工作區歷史與高水位，並在有保護後建立遷移紀錄，測試並行、失敗與中斷。未知歷史不能直接從 01 開始。
- Actions、集中配號程式、測試及實際遠端驗證均未完成。此前工作目錄既有修改保留在 index；本 Issue 的草案尚未提交或推送。

## 2026-11-17 確認與狀態
使用者確認日期序號歷史完整，並建立 ruleset 27431700；查詢確認目標、狀態、禁止刪改及無 bypass。已建立歷史參照 20261117/01。實作及自動測試進行中，真實 GitHub Actions 執行、artifact 與維運就緒仍待驗證。原有本機預檢環境衝突以獨立測試版本 fixture 驗證流程，未更動主機 COM。

# ISSUE-0008｜GitHub Actions 自動打包需求草案 v0.1

狀態：使用者已核准 v0.1；PM：Codex。類型：變更；等級：L3，因跨 GitHub 與本機的系統整合、權限及安裝器版本識別契約。目標：由 GitHub Actions 自動建置 ExcelKit Inno Setup 測試安裝包，跨次執行避免版本重複。

## 範圍與驗收
- Windows GitHub-hosted runner 取得特定提交，依 Installer/Toolchain.json 安裝精確版本 .NET SDK 與 Inno Setup，核驗版本並使用鎖定 NuGet 套件及既有打包程序；不可假設 runner 預裝版本固定。
- 產出 EXE、manifest、build report，綁定提交 SHA、執行識別與雜湊，作為 Actions artifact；不建立 Release、不部署、不更動正式版基準。artifact 保存期限及使用權限以實際儲存庫設定為準。
- 沿用 `<DLL FileVersion>-yyyyMMdd-NN` 格式及台北日期，NN 至少兩位。打包前以唯一共同分配來源保留序號，並行與重試不可重用；失敗、取消亦保留。不能確定唯一性就停止。移轉既有保留紀錄與歷史產物，含部分失敗紀錄；無法完整盤點時停止。
- 本機可繼續打包，但須接同一分配機制；不能連線或授權不足時停止，不退回本機獨立配號。保留已存在的輸出及保留紀錄，不覆蓋。
- 軟體版本、安裝器檔名、清單、記錄與產物一致，排除現行指定影片；對應驗證、失敗訊息及操作說明。雲端建置成功不代表安裝、COM、LabVIEW 或互動 UI 驗收。

## 適用性與邊界
架構：適用，新增 GitHub 打包與分配邊界；UI：不適用，無 UI 設計；資料庫：不適用，不新增或遷移應用資料庫；系統整合：適用，GitHub API、runner 與本機協調；維運：適用，永久保留序號、失敗恢復、權限與稽核。Codex 依序執行 PM、SA、系統整合及測試規劃，非獨立審查。

來源：src/Excelopen；安裝器：Installer；本機輸出：publish/Installer/Setup；雲端輸出：workflow artifact。DLL 產品版本、FileVersion、AssemblyVersion 與正式基準沿用既有規範；涉及 CI 的新輸出版本策略需在實作前確認一致。SA 方案見 Architecture/AR-ISSUE-0008.md，測試草案見 TestPlans/TP-ISSUE-0008.md。方案已核准並進入實作，測試結果見 TestReports/TR-ISSUE-0008.md。

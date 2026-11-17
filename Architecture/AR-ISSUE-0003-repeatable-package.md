# 可重複打包與安裝驗證架構評估｜ISSUE-0003

Codex 依序扮演 PM、SA、系統整合、維運、軟體及測試，屬自查；2026-11-16。需求與人類核准：Requirements/RS-ISSUE-0003-repeatable-package-draft.md；測試：TestPlans/TP-ISSUE-0003-repeatable-package.md。既有分支 feature/ISSUE-0003-installer，固定測試輸出 publish/Installer/Setup，正式基準 1.2.4。

## 現況、方案與取捨

現況 MSBuild Release 會複製到既有 publish，且不清除舊檔；Config/Rebuild-Setup.ps1 原先直接依賴該處。建議並實作固定腳本讀取明確來源，建立一次性空白 staging，在其中使用鎖定 NuGet、.NET SDK 及 Inno Setup 建置與核對，再不覆寫地產生測試安裝檔、安裝清單與建置報告。相較於重用 publish，增加暫存空間與執行時間，但可避免舊檔混入並保留失敗證據。程式資料夾 Installer，工具鏈由 Installer/Toolchain.json 指定，元件套件鎖於 src/Excelopen/packages.lock.json；正式基準由 FormalBaseline.txt 管理。

```mermaid
flowchart LR
  S[來源與鎖定設定] --> P[預檢與獨立快照]
  P --> B[兩框架建置與版本核對]
  B --> M[安裝清單與中繼資料]
  M --> I[Inno 編譯]
  I --> O[固定 Setup 目錄：exe、清單、報告]
  O --> V[安裝後唯讀驗證]
```

應用程式資料庫與 UI 變更不適用；安裝器程式新增安裝選項狀態檔、內含版本清單及驗證腳本。系統整合適用，需確認 Windows、.NET、Inno、x86/x64 COM 與登錄視圖。維運適用，需保留 staging、清單、雜湊及失敗原因；正式基準變更須另依專案核准流程。此流程不自動註冊、安裝、部署、提交或合併。

## 風險與回復

- 輸入清單記錄來源與成品 SHA-256，鎖定依賴與工具版本；建置時間等仍可能使產物位元組不同。清單與選項檔本身不是數位簽章；驗證器可接受外部受信任的 manifest SHA-256，未提供時標記來源驗證未執行。
- 產出僅以不覆寫方式移到固定位置。若發布途中發生磁碟或程序故障，可能留下部分新檔；不得把不完整集合視為成功，保留 staging 與既有檔案，人工核對日誌及雜湊後處理；不可自動遞迴清除。
- 安裝選項紀錄儲存在安裝程式管理目錄；驗證器也檢查交易狀態必須完成。失敗、遺失、拒絕讀取、使用者層覆寫、未知 COM 版本與不一致皆不得當成通過。先前升級遺留的未選可選檔不作本次錯誤判斷。
- 唯讀檢查不改變安裝或登錄；明確選用 Functional 時會建立 COM 物件並以記憶體工作簿測試，必要時可能產生日誌，不應自動對正式環境執行。
- 安裝器新增選項紀錄、COM 成功交易、升級保留、解除安裝、權限拒絕、雙視圖及故障回復需隔離 Windows 實測，未完成前整體架構符合性及維運就緒未確認。

## 結論與實作後檢查

有條件採用；使用者已確認草案。版本 1.2.4+20261116141740 已編譯並通過 TestReports/TR-ISSUE-0003-repeatable-package.md 記載的預檢、快照、雜湊與隔離樣本測試。真實安裝及 COM 功能／回復未執行；相關必要驗證仍阻擋整體驗收及合併建議。
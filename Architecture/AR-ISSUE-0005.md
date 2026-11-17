# ISSUE-0005 架構評估草案 v0.1

SA 與系統整合：Codex，同一執行者自查，非獨立審查。依據現有工作區 `src/Excelopen/Excelopen.csproj`、`Installer/Package-ExcelKit.ps1`、`Installer/ComTransaction.ps1`、`Installer/Startup.iss`、`Installer/Verify-ExcelKit.ps1`、留存 build 報告；未提交狀態。關聯 Requirements/RS-ISSUE-0005.md。狀態：待核准。

## 現況與提案
現況：COM 客戶端 → 已註冊 IDispatch / .NET Assembly `1.2.4.0` → ExcelKit DLL（檔案版本 `1.2.4.0`）；打包與驗證分別檢查 assembly、file、product version。安裝器的檔案版本顯示來自已登錄 DLL，並以 hash 比較套件。已留存的 5 個報告 FileVersion 均 `1.2.4.0`，無法證明其他機器的歷史。

提案：沿用 COM 介面和 AssemblyVersion，將檔案版本改為正式基準加單調遞增 revision，納入受控版本分配紀錄與打包檢查；保留產品版本時間戳，並核對產出 DLL。版本分配要在編譯前完成；一旦分配不回收，以免失敗重試與其他版本重用。多分支／多主機必須先集中協調版本，無集中一致記錄時不保證全球唯一。本次維持既有安裝畫面，DLL 檔案屬性及打包報告可供查詢。

## 方案比較
| 方案 | 相容及可操作性 | 風險 |
| --- | --- | --- |
| A：維持組件版本、遞增 FileVersion | COM 註冊識別不隨測試版變動；Windows 檔案版本可比較 | 需集中管理序號；COM 動態與升級驗證仍必要 |
| B：同步遞增 AssemblyVersion 與 FileVersion | 組件及檔案皆有新編號 | 現有註冊偵測、舊客戶端與回復可能受影響，需要更大範圍遷移 |
| C：只改產品版本 | 維持既有實作 | 無法滿足使用者要求 |

建議 A，有條件採用；不新增技術棧。架構符合性檢查：待執行。

## 測試與回復
比較舊版與新版本的 interface IID、CLSID、ProgID、型別庫版本、所有舊 DispId 與參數簽章，並於隔離環境測試 x86/x64 實際 COM activation、舊 API、新功能、安裝升級及回復；讀取 DLL FileVersion、產品版本與 AssemblyVersion。檢查重複與倒退版本能被拒絕，版號上限及多來源衝突能安全停止。先前 Sandbox 出現 Application Control 阻擋，必須實際確認；受阻時不宣稱 COM 相容並依使用者要求停止交付打包。回退以隔離環境保留的舊安裝包與登錄快照測試；不得推定已驗證。無新連網、機密或常駐監控；效能目標未提出。

## 決策與待決
人類需核准需求及架構草案 v0.1；UI 不適用，PM 須完成測試計畫後始可開發。實作後 SA 符合性與 COM 測試均待執行；不以靜態檢查取代執行證據。

## 實作後自查（2026-11-16）
使用者確認保留既有 ExcelKit.ExcelOperator；候選版本組件 1.2.4.0、檔案 1.2.4.1、產品 1.2.4+20261116223400。匯出型別庫二進位一致，CLR COM 身分及簽章一致；版本台帳及預檢已實作。Sandbox Application Control 阻擋舊安裝器與 RegAsm 載入 DLL，動態相容、升級及回復未驗證；架構符合性尚未確認。詳細見 TestReports/TR-ISSUE-0005.md。

## 實作後符合性複核（2026-11-16）
Codex 自查：兩版靜態 COM 介面一致；實際候選安裝器於 Windows 11 x64 Sandbox 完成升級、x86/x64 COM、manifest 驗證及舊版回復，均通過，證據見 TestReports/TR-ISSUE-0005.md。FileVersion 1.2.4.1、AssemblyVersion 1.2.4.0，符合已核准方案；未覆蓋舊版安裝包。LabVIEW、32 位元 Windows 與互動 GUI 未執行，不推論其通過。版本台帳需在不同工作區及主機間集中協調，無法自動查得所有歷史版本。

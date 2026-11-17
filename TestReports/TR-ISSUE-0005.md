# ISSUE-0005 檔案版本與 COM 驗證結果

2026-11-16，Codex 依序執行 PM、SA、軟體、測試，自查非獨立審查。基礎 main 1ad27cb8a04e6a66f022e9fda164c999e8fafd26；feature/ISSUE-0005-file-version；含 ISSUE-0004 與其他未提交工作。需求與架構 v0.1 經使用者確認；使用者另確認保留 ExcelKit.ExcelOperator。

## 實作
檔案版本 1.2.4.1，產品版本 1.2.4+20261116223400，AssemblyVersion 1.2.4.0。Installer/FileVersion.ps1 與 Installer/FileVersionLedger.json 建立保留及驗證規則；預檢核對目前來源摘要、台帳、可見 build 報告與本機 COM DLL，防止版本重用／倒退與序號越界。以同機、同路徑互斥鎖控制並行，跨機或不同工作區須集中協調台帳；無法自動探知未納入的歷史版本。版本一經保留不回收。來源變更須新產品時間戳與較高序號。主機已安裝 DLL 唯讀檢查為 1.2.4.0，未變更其登錄或檔案。

## 結果
- 通過：`dotnet build src/Excelopen/Excelopen.csproj -c Release --no-restore -p:PublishToFolder=false -v:q`，net472/net10.0 建置成功，零錯誤，既有 23 個 nullable 警告。兩份 DLL 的 FileVersion=1.2.4.1、AssemblyVersion=1.2.4.0、產品版本含 1.2.4+20261116223400。
- 通過：版本台帳 Check 及唯讀打包預檢；隔離副本測試來源變動拒用、分配 1.2.4.2、重複配置拒絕、低版本拒絕、版本超限拒絕、既有報告衝突拒絕；未改來源可重複 Check。初版台帳更新的 Windows PowerShell File.Replace 問題已修正並重測；初版打包腳本 UTF-8 編碼問題已修正並重測。隔離測試中的還原比較曾因檔案編碼變更引發摘要失配，屬防護正確拒絕，不計成功重建，改以實際未改來源重複檢查。
- 通過：使用相同的 TlbExp 匯出前版 1.2.4.0 與新版 1.2.4.1 的 net472 DLL 型別庫，兩份 SHA-256 均 55A9E905A97B006124FF97F08B6C06858D3475DD2B799A04710AF3AFA0330E44；舊／新 CLR 的 ProgID、CLSID、IID、AssemblyVersion、型別庫 GUID、全部 28 個公開方法與 DispId／參數簽章比較一致。這是靜態證據，不代表 COM 註冊與執行已通過。
- 通過：ISSUE-0004 功能測試與顏色範例以新版 Release DLL 執行；以 openpyxl 檢查重開檔案的 RGB。手冊 MD、HTML、PDF 補充檔案版本與穩定 COM 名稱；檢視 PDF 版面。
- 未執行／受阻：Windows Sandbox 中舊安裝器遭 Application Control policy 阻擋（`.tools/installer-tests/issue-0005/results/setup-result.txt`）；獨立直接 RegAsm 嘗試載入舊 DLL 亦遭阻擋（`results/result.txt`，HRESULT 0x800711C7）。因此 x86/x64 實際 COM activation、升級、回復、安裝與 LabVIEW 測試未通過也未完成。另試無登錄 activation context，但 class factory 回報 0x80070002；不得據此推論已驗證 COM。Sandbox 的政策沒有被繞過。
- 未執行：依使用者要求先確認 COM 相容性再打包，故尚未建立 1.2.4+20261116223400 安裝包。執行過 Preflight，未發布或安裝；保留舊測試包。無主機 COM 註冊或安裝變更。

結論：檔案版本與靜態檢查通過；動態 COM 及安裝／升級回復驗證受阻，整體未結案，待可執行測試的隔離 Windows 環境後繼續。

## 2026-11-16 Sandbox 再次診斷
Windows Sandbox 可啟動，網路停用；再次嘗試舊測試安裝器與舊 DLL，均受阻。Sandbox `Microsoft-Windows-CodeIntegrity/Operational` 事件 3077、3033 記錄其未滿足簽署等級或違反政策，policy ID `{0283ac0f-fff1-49ae-ada1-8a933130cad6}`，並有 Smart App Control Block Details 事件。主機唯讀 AuthenticodeSignature 檢查兩個舊測試輸入均為 NotSigned；不能推定簽章必定解決或將失敗原因歸於單一因素。完整隔離事件見 `.tools/installer-tests/issue-0005-diagnostic/results/Microsoft-Windows-CodeIntegrity-Operational.txt`，其他輸出與診斷腳本位於同次測試目錄。沒有修改安全原則。COM 執行、安裝、升級及回復仍未驗證，依使用者要求未重新打包。

## 2026-11-16 受託調整 Sandbox 後續
使用者明確授權僅在 Sandbox 調整設定後重跑相容性測試。已建立獨立測試設定 `.tools/installer-tests/issue-0005-authorized/manual.wsb`，網路停用、輸入唯讀及單獨輸出，並修正測試腳本：若舊安裝器失敗即回報失敗而不繼續誤判。電腦操作後端兩次啟動失敗，均為 `windows sandbox failed: helper_unknown_error: setup refresh had errors`，因此未取得 Sandbox 介面且未更動任何 Smart App Control 或主機設定。尚未執行本輪 COM 實測或重新打包。

## 2026-11-16 使用者執行測試與後續更正
使用者在獨立 Sandbox 執行 `C:\SandboxInput\Test.ps1`，結果 `FAIL old install`（`.tools/installer-tests/issue-0005-authorized/results/`）；之後使用者更正，尚未在 Sandbox 關閉智慧型應用程式控制。故該結果不能視為關閉後之測試，亦不能證明設定變更有效。未執行 COM 階段、未打包；待設定完成再測。電腦操作後端目前仍無法初始化。

## 使用者提供的 Sandbox 診斷
使用者於同一 Sandbox 執行唯讀事件查詢；`issue-0005-authorized/results/policy-diagnostic.txt` 含 2026-11-16 23:11:30 的 3118、3077、3033 事件，顯示測試安裝器受 policy `{0283ac0f-fff1-49ae-ada1-8a933130cad6}` 與簽署要求阻擋。該紀錄是先前測試時的歷史事件，不證明執行查詢當下的原則狀態。使用者先前澄清未關閉智慧型應用程式控制。未變更主機或 Sandbox 原則，動態 COM 測試仍受阻。

## 2026-11-16 Sandbox 設定調整後測試
使用者在 Sandbox 內依授權調整測試用原則後，重跑測試。`issue-0005-authorized/results/old-setup.log` 與 `checks.txt` 顯示舊測試版安裝成功；舊 `setup-result.txt` 為先前 23:11 的失敗紀錄，非此次結果。其後 `result.txt` 顯示 RegAsm `/codebase` 的未強式名稱簽署警告被 PowerShell 的 `$ErrorActionPreference=Stop` 當成 NativeCommandError，測試腳本提前中止；尚未完成 COM activation，不代表實際 COM 不相容。需修正測試腳本對 native stderr 的處理，並依程式退出碼重測 x86/x64 與回復，必要測試未通過前不打包。

## 2026-11-16 Sandbox COM 重測及打包
使用者在同一 Sandbox 依已授權設定後重跑修正的測試；`issue-0005-authorized/results/result-v2.txt` 為 PASS，checks-v2 紀錄 old x86、old x64、new x86、new x64、old x86、old x64 各通過。四份保存活頁簿以 openpyxl 重開，確認舊 API 資料與紅／藍 RGB。舊版安裝成功。此處的切回舊版是使用 RegAsm 重註冊，不等於安裝器升級及回復測試。

先測試的新版 net472 DLL SHA-256 `BFC88C57EE9FF7F0E613236D25159B673D7379A2705E7289750518C58423ABDA`。COM 測試後以 `Package-ExcelKit.ps1 -Preflight -InstallerVersion 1.2.4+20261116223400` 及同版打包成功；安裝器 SHA-256 `5EFCFBB8A413AF84153DA5E1715DBADBD7D66D003FB655DFAA2800A14AB936B9`；manifest SHA-256 `22B09E694547DBB518B74223B8C8AC6EA99660322C003A2553B2F6F504C03771`。比對 26 個 payload、30 個來源與打包報告通過；`Installer/Tests/Test-Package-Workflow.ps1` 的必要清單、竄改、選項及 traversal 測試通過。

打包流程另行建置的 net472 DLL SHA-256 `618D526761667591F1C00E10306958CDDB491F809E49D0668A16C76854289094` 與先前直接測試的 DLL 不同，不能將先前動態測試直接認定為包內 DLL 通過。已準備 `C:\SandboxOutput\package-test-20261116223400\Run-Package-Test.ps1`，以同一 Sandbox 實際測試安裝包升級、x86/x64 COM、安裝清單驗證及舊包回復。該測試尚未執行，安裝包目前為待驗證候選，不得宣稱整體通過。

## 2026-11-16 實際安裝包升級及回復驗證
使用者在已調整測試設定的 Windows Sandbox 執行 `package-test-20261116223400/Run-Package-Test.ps1`；`result.txt` 為 PASS，證據位於 `.tools/installer-tests/issue-0005-authorized/results/package-test-20261116223400/`。

- 通過：候選安裝器 SHA-256 與建置報告一致；舊版安裝 DLL hash 核對後，恢復 x86/x64 已安裝路徑的 COM 註冊；新版安裝器升級 exit 0；安裝後 net472 DLL SHA-256 為 `618D526761667591F1C00E10306958CDDB491F809E49D0668A16C76854289094`，與本次安裝包清單一致，FileVersion `1.2.4.1`。
- 通過：實際安裝新版後，x86、x64 以 `ExcelKit.ExcelOperator` 呼叫既有讀寫方法和文字顏色方法，均回報產品版本 `1.2.4+20261116223400`。同版 manifest SHA-256 交給 Verify-ExcelKit.ps1 加 Functional 驗證，overall pass、12 個 pass。
- 通過：以先前測試安裝器回復，exit 0；DLL hash 恢復為 `30245AB1D97D3497ED950643AE1FF02BB6248D919CBA954E3EC47A91DA907750`，FileVersion `1.2.4.0`；x86/x64 COM 與舊版 manifest 加 Functional 驗證 overall pass、12 個 pass。
- 通過：測試產生的 `new-x86.xlsx`、`new-x64.xlsx`、`old-x86.xlsx`、`old-x64.xlsx` 以 openpyxl 重新開啟，A1 文字與紅色、B2 藍色正確。
- 範圍：這是 Windows 11 x64 Sandbox 內的自動化安裝及 COM 測試，並非 LabVIEW 實際整合、32 位元 Windows 或互動安裝 UI 驗收；這些未執行。測試後 Sandbox 處於舊版回復狀態，不代表已更新主機。

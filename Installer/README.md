# ExcelKit 安裝程式與重新打包

測試版安裝器版本由打包腳本指定；正式版基準記錄於 `FormalBaseline.txt`。固定輸出位置為 `publish/Installer/Setup/`。工具版本由 `Toolchain.json` 鎖定；安裝器版本與 DLL 產品、組件及檔案版本分開記錄。

## 安裝方式
安裝與解除安裝會要求 Windows 管理員權限。提供繁體中文、英文，預設依 Windows 顯示語言，可自行切換。安裝時選擇供 x86、x64 或兩種客戶端使用的 COM；升級既有安裝時需保留已安裝的位元數。操作前請關閉使用 ExcelKit 的程式。

另可獨立選擇操作手冊（含媒體）、範例程式及 .NET 10.0 元件；首次安裝時手冊及媒體、範例與 .NET 10.0 均預設不勾選。三項均不選時仍會安裝及註冊 COM 必要的 .NET Framework 元件。重新安裝取消選項不會主動移除先前安裝的檔案；可透過解除安裝移除安裝器管理的檔案。重新安裝可能沿用上次的任務選擇。

若偵測到可辨識的手動註冊 ExcelKit 1.2.0.0，會顯示舊版的位元、組件及來源，詢問是否解除舊註冊後安裝新版。預設不同意；只有同意才進行。辨識包含舊 DLL SHA-256 `B2C2FC06817EA615CED43C90D910F259E516E4A10BEE3941DC701EB1339FA84E`，其他未知版本、遺失 DLL、使用者層覆寫、多版本或不一致註冊會停止並保留原狀。

舊 DLL 及外部資料夾保留；成功移轉後其他程式新建立的 COM 物件會使用新版。解除新版安裝不會自動重新註冊舊版。

## 備份與失敗回復
變更前，在 `%ProgramData%\ExcelKit\Transactions\<交易編號>` 保存兩種位元視圖的產品登錄樹（ProgID、CLSID、Interface、TypeLib）、值型態、DACL 及安裝擁有者資料。備份有 SHA-256 驗證，資料夾僅管理員和 SYSTEM 可寫；同一安裝升級另備份既有 net472 元件檔。

移除舊版時使用 `/unregister` 並清理已驗證的產品中繼資料，以支援舊 TypeLib 暫存檔已遺失的情況。若移除或新版註冊失敗，回復原登錄並核對內容；已安裝版本升級失敗時也還原原元件。失敗結束碼非零，完成頁顯示失敗；可能保留未啟用的新安裝檔案，可執行其解除安裝，或重新執行安裝器重試。備份及 `error.txt`、`regasm.log` 保留供診斷。

若程序被強制終止或斷電，下一次操作會阻止覆寫未完成交易。由管理員依錯誤中的交易路徑回復（先確認交易屬於此安裝）：

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ComTransaction.ps1 -Action Rollback -AppDir "C:\Program Files (x86)\ExcelKit" -Architectures x86,x64 -Transaction "C:\ProgramData\ExcelKit\Transactions\<交易編號>"

已完成交易拒絕用此命令回滾。機器被強制終止的完整檔案安裝流程，與一般可捕捉的註冊失敗不同；請依記錄處理，勿自行刪除登錄樹。

## 重新打包

安裝器版本採 `<DLL FileVersion>-yyyyMMdd-NN`，日期為台北時間，同日從 01 依序遞增；檔名 `ExcelKit-Setup-<安裝器版本>.exe` 不含 TEST，仍屬測試版。指定影片 `Reference/Media/Video_2026-08-15_233419.mp4` 不納入新安裝包，來源與舊包保留。已有版本、失敗保留號與並行衝突不可覆寫；本機與 GitHub Actions 使用受保護的 GitHub 參照集中保留序號。DLL 產品版本與 COM 組件版本維持既有規則。

從專案根目錄使用同一個唯一測試版號先預檢再打包，或透過專案 skill `excelkit-package` 執行：

```powershell
powershell.exe -NoProfile -NonInteractive -File .\Installer\Package-ExcelKit.ps1 -Preflight
$version = '<預檢輸出的 Installer 版本>' # 例如 1.2.4.3-20261117-01
powershell.exe -NoProfile -NonInteractive -File .\Installer\Package-ExcelKit.ps1 -InstallerVersion $version
```

預設使用 `.tools/inno-portable/ISCC.exe`；如在其他位置，兩次均可指定 `-CompilerPath <ISCC.exe>`。腳本使用獨立 staging、鎖定 NuGet 和工具鏈，核對 DLL、素材及版本，輸出 `.exe`、`.manifest.json`、`.build.json` 至 `publish/Installer/Setup/`。同版不覆寫；失敗時保留 `.tools/package-work/` 供診斷。舊 `Rebuild-Setup.ps1` 已改為呼叫標準腳本，不再從既有 `publish` 殘留檔直接打包。

## 安裝後驗證

安裝器會保存在 `ExcelKit\Installer` 的驗證腳本、清單與安裝選項紀錄。建議從**同一版本的建置報告**取得 `manifest_sha256`，再執行：

```powershell
powershell.exe -NoProfile -NonInteractive -File .\Installer\Verify-ExcelKit.ps1 `
  -InstallDir '<ExcelKit 安裝資料夾>' -ExpectedManifestSha256 '<同版 manifest_sha256>'
```

預設唯讀比對檔案、版本、選取內容、COM 註冊與交易結果，輸出 JSON；`overall` 為 pass／fail／unknown，結束碼分別為 0／1／2。未提供外部清單雜湊時，清單來源驗證標為 `not_run`。只有使用者要求功能檢查或在已授權的隔離環境，才加 `-Functional`：它會啟動 COM 並以記憶體工作簿測試，可能啟動程序或產生日誌。此腳本不修復、不重新註冊，不能取代 GUI、升級、列印及其他目標環境驗收。

## 驗證範圍
Windows 11 x64 Sandbox 的 x86／x64 客戶端測試及故障注入結果請見專案 `TestReports/TR-ISSUE-0003.md`。32 位元 Windows 本身、LabVIEW、LibreOffice PDF／列印及強制斷電未實測。本次是開發測試版，沒有改變正式版基準。

本次另含三格式手冊、VBS／LabVIEW／PowerShell 範例與開始功能表資料夾入口，檔案及範例驗證見 TestReports/TR-ISSUE-0003-examples.md。


安裝完成頁可選擇開啟 ExcelKit 安裝資料夾、LibreOffice 安裝版或免安裝版下載頁，三項預設不勾選；靜默安裝不開啟瀏覽器。安裝資料夾的 `LibreOffice-下載連結.txt` 保存下載頁網址。

歡迎頁顯示已註冊 DLL 及套件版本資訊，並可直接開啟下載頁或已確認的安裝資料夾，不需繼續安裝。版本偵測僅讀取登錄與 DLL，註冊資訊不等同搜尋整台電腦；無法確認時顯示未知。啟動時仍可能要求管理員權限，實際安裝前會檢查 .NET Framework。

## GitHub Actions 與集中序號
推送 main 或 feature/* 分支，以及手動執行 workflow 時，GitHub Windows runner 會安裝鎖定版本工具、執行配號及預檢測試並打包。成功時在該次 Actions 的 artifacts 下載 EXE、manifest 與 build report，並建立附有相同三份檔案的 GitHub 測試版 Release；不自動部署。報告記錄 source_commit、GitHub run ID、attempt 與檔案雜湊。artifact 可能依 GitHub 設定過期；請依需要另行保存。

新版打包須使用原 GitHub origin 和網路；透過受保護的 `refs/tags/excelkit-package/<台北日期>/<序號>` 原子建立紀錄。本機需 GitHub push 權限，CI 使用 workflow 的 contents:write；任何失敗或取消不回收已建參照。歷史 20261117-01 已匯入，該歷史參照指向導入時基底而非舊包的來源 commit，舊包資料請看原 build report。不可刪改參照或停用保護規則；配號前驗證 ruleset ID 及已審核更新時間。規則改動會使打包停止，需管理員重新核對規則並更新受信任基準。管理員仍可能修改規則本身，應定期檢查設定與日誌。

本機保留序號前需先提交 src 與 Installer 的變更；未追蹤的 src 亦會阻擋。只有 GitHub 保護及既有序號記錄可查時才能打包，不可使用舊版本的本機獨立配號流程；離線、權限不足、無法確認推送結果或衝突時停止，先檢查遠端參照後重新預檢。跨工作區不得使用舊腳本自行配號。雲端建置與靜態檢查不等於安裝、升級、COM、GUI 或 LabVIEW 的驗收。
## 下載 Setup 與 JSON 的用途
在 GitHub Releases 選擇標示「測試版」的版本，在 Assets 點選 `ExcelKit-Setup-<版本>.exe` 安裝；GitHub 顯示的 Source code ZIP/TAR 是原始碼，不是安裝程式。安裝一般只需 EXE，兩份 JSON 可供檢查與追溯：
- `.manifest.json` 列出安裝內容及各檔 SHA-256，供 Verify-ExcelKit.ps1 核對；安裝器本身亦攜帶安裝所需清單。
- `.build.json` 記錄來源提交、工具版本、輸入與安裝包雜湊等建置資訊，用於追溯及驗證。
公開前 workflow 先建立 draft，下載核對三個 Release 資產後轉成 prerelease。若任何步驟中斷，保留 draft 並停止；遇到同版 Release 不覆寫，請由維運人員檢查 GitHub，再依情況處理，不重用序號。測試版不代表已驗證實際安裝、升級或 COM。
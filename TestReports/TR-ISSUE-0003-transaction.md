# 最終安裝器測試結果｜ISSUE-0003

日期：2026-11-15（Asia/Taipei）；執行者：Codex（軟體、SA、維運及測試，自查，非獨立審查）。分支 `feature/ISSUE-0003-installer`；未提交，保留先前混合的未提交工作，未合併或推送。

## 交付版本及位置
正式版基準 1.2.4；交付為測試版 `1.2.4+20261115075927`。來源 `src/Excelopen/`，固定元件輸出 `publish/bin/`，安裝程式 `publish/Installer/Setup/ExcelKit-Setup-1.2.4-20261115075927-TEST.exe`；設定、COM 交易助手及說明在 `publish/Installer/Config/`。

安裝檔 SHA-256：`9E11F6BD6BCAC6298A2A139D23D2181C440B97A629CDCA6BA0D00A873B198D03`。
助手 SHA-256：`E9325EF33ECEA7E2CD3C4CFA735387BF37695CF6FAEBA98D4C79AB69FF26E0D8`。
完整清單：`publish/Installer/Config/release-manifest.json`。

## 已完成變更
安裝器辨識已確認的手動舊版 1.2.0.0（以實際 DLL 指紋及 COM 身分核對），顯示來源與位元並詢問同意後替換；保留舊 DLL。交易助手備份固定產品登錄樹及 DACL、驗證備份雜湊，支援舊 TypeLib 檔案遺失，註冊失敗時回復登錄；同目錄升級亦備份及還原原 net472 檔案。未完成交易、備份損壞、未知 DLL 和使用者層覆寫會擋下。未知版本不自動移除。成功移轉後解除新版安裝不會自動重新註冊舊版。

## 環境及命令
Inno Setup 7.1.1 免安裝編譯器；Windows 11 x64 Sandbox（10.0.26100，Windows PowerShell 5.1／.NET CLR 4）。停用網路，輸入唯讀，專用結果資料夾可寫。產品候選與測試安裝器使用相同來源，故障參數僅納入以 `/DTestBuild=1` 編譯的 Sandbox 測試版，交付安裝器未編入故障參數傳遞。COM 測試由 x86、x64 Windows PowerShell 實際 CreateObject、讀取版本及活頁簿儲存格寫入／讀回；Sandbox 沒有 VBScript 引擎，故未使用 VBS。

建置：`dotnet build src/Excelopen.slnx -c Release --no-restore`；0 錯誤，23 項既有 net10 可空性警告。安裝器：`ISCC.exe publish/Installer/Config/ExcelKit.iss /Qp` 成功。設定及助手來源與 Config 副本雜湊一致。

## 通過的產品整合案例
證據目錄：`.tools/installer-tests/sandbox/results/integration-20261115-082118/`，`result.txt=PASS`。

- x86、x64、雙位元新安裝；各適用位元 COM 實際呼叫及版本一致。
- 三種組合解除安裝成功，指定產品 COM 登錄恢復乾淨基準。
- 拒絕舊版移轉：退出碼 7，原登錄保留。
- 舊 TypeLib 檔不存在時，移除後故障與新 x86 註冊後故障：退出碼 1，回復原登錄；Inno 日誌明確記錄 Rollback exit=0。
- 兩種故障安裝的解除清理：不動到回復的舊版註冊。
- 舊版雙位元成功移轉、保留舊 DLL、重新安裝及雙位元 COM 呼叫。
- 解除安裝故障：退出碼 1，登錄回復、DLL 保留、兩種 COM 呼叫仍通過；重試解除安裝成功且登錄清乾淨。

## 通過的保護及版本升級案例
`.tools/installer-tests/sandbox/results/safety-20261115-082720/`，`result.txt=PASS`：備份建立、未完成交易攔截、損壞備份於變更前攔截、回復準備中的交易、未知註冊與使用者層覆寫不被更動。

`.tools/installer-tests/sandbox/results/upgrade-20261115-082747/`，`result.txt=PASS`：從先前安裝版升級時，在 x86 註冊後注入失敗，原登錄及原 DLL SHA-256 都恢復；後續成功升級、雙位元 COM 呼叫及解除安裝通過。

## 未執行及界線
純 32 位元 Windows、LabVIEW、LibreOffice PDF／列印、互動精靈視覺驗收、強制斷電和任意登錄並行修改未測。測試的同意／拒絕採明確的靜默參數，標準 GUI 確認對話框未做自動點擊驗收。回復驗證範圍為產品資料、值型態及 DACL，不宣稱完整系統快照或登錄擁有者／稽核規則的原樣還原。未知舊版持續保守停止。

## 角色結論
SA：實作符合已核准的受控舊版取代及交易回復方向，以上界線保留。UI：繁中／英文與位元選擇已保留；GUI 視覺未驗收。DBA：不適用，無應用資料庫變更。系統整合：Sandbox 的雙位元 COM 已驗證。維運：受保護備份、診斷日誌與未完成交易的回復操作已記錄於 Config/README.md；不是整機復原機制。

主機最後唯讀核對時已發現新版 COM 指向 `C:\Program Files (x86)\ExcelKit`，安裝助手雜湊與本次交付一致；無法只由這項觀察確認是誰或何時安裝，因此不宣稱主機仍維持舊版註冊。

使用者於本次對話確認已在測試期間自行執行新版安裝；這與主機唯讀核對的新版路徑及助手雜湊一致。

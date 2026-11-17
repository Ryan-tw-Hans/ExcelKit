# LibreOffice 整合可行性檢查｜ISSUE-0003

執行者：Codex（SA、系統整合、測試工程師，非獨立審查）；2026-11-16。基準：使用者確認之 RS-ISSUE-0003-libreoffice-draft.md。

## 已驗證

- 來源：LibreOffice 官方連結之 PortableApps.com 26.2.4 MultilingualStandard，下載檔 .tools/downloads/LibreOfficePortable_26.2.4_MultilingualStandard.paf.exe，223657592 位元組。SHA-256：4BDE93374AEF4409243505B20D16561A4628AC7591457DD01FC6E1CCF571BA65，符合發行者公布值。
- 以 7-Zip 解壓到 .tools/LibreOfficePortable-26.2.4，完成，共 15066 檔、927197182 位元組。原廠封裝程式以 /S 嘗試解壓但停在互動視窗，已停止本次程序；未判定原因。
- App/AppInfo/Launcher/LibreOfficePortable.ini 的 FileWrite1、2、13–16 寫入 App/libreoffice/program/bootstrap.ini 與 setup.ini；UserInstallation 指向套件內 Data/settings。現有 Installer/ExcelKit.iss 預設安裝位置為受保護的 Program Files，且一般使用者通常不能修改該位置。直接在此目錄執行 portable launcher 因此有權限衝突風險；尚未實機驗證。
- src/Excelopen/Class1.cs 的 GetLibreOfficePath 只檢查明確指定路徑及系統 LibreOffice 位置／PATH；不會自動尋找本套件。明確 SetLibreOfficePath 可指定路徑，但實際 PDF／列印仍未驗證。
- 發行者文件指出非 ASCII 安裝路徑可能使程式無法正常運作；本機專案路徑包含中文。另需核實以 7-Zip 提取是否等同廠商安裝程序的輸出與設定。

## 結論與替代方案

既定方案以 {app} 子目錄保存套件，要求標準使用者可執行並保存設定；上列證據顯示直接嵌入並指向原廠 launcher 可能不滿足此條件。按已核准草案「若無法安全滿足則停止並回報」暫停修改安裝器；不可將靜態推論當實測通過。

建議進一步評估：程式集中保存在受保護位置，使用專用啟動方式將使用者設定放在每位使用者可寫位置，或首次使用時複製完整 portable 套件到每位使用者的本機資料夾。前者需驗證不使用原廠 launcher 時的初始化及可攜性；後者會增加每位使用者磁碟用量，並需定義升級及資料保留。均須驗證中文路徑、既有 LibreOffice 共存、ExcelKit 呼叫、安裝／卸載與授權；尚未確認可行，不可視為已核准的最終設計。

建置、安裝、GUI、標準使用者、PDF、列印、升級、卸載、COM 回歸與故障回復：未執行。測試計畫在 TestPlans/TP-ISSUE-0003-libreoffice.md。既有正式版基準 1.2.4 未變，沒有產生新版安裝程式。
## 直接啟動探測

以解壓後的 App/libreoffice/program/soffice.exe 執行 --headless --version，30 秒未結束，已停止本次啟動的程序；原因未確定。此結果不能證明可攜版可用，也不能證明受 Application Control 阻擋。替代啟動方式及一般使用者測試仍未完成。

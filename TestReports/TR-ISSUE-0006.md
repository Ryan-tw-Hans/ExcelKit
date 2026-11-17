# ISSUE-0006 測試進度

2026-11-17；Codex 依序擔任 PM／設計／開發／測試，為自查而非獨立審查。需求與設計 v0.1 已由使用者確認。候選元件產品版本 1.2.4+20261117065015，檔案 1.2.4.2，組件 1.2.4.0，正式基準 1.2.4。

- 通過：Release 建置 net472/net10.0，0 錯誤，23 項既有 nullable 警告；net472 DLL FileVersion 1.2.4.2 與 AssemblyVersion 1.2.4.0。
- 通過：在 `.tools/installer-tests/issue-0006-compile` 隔離編譯測試用安裝畫面；以實際 net472 DLL 生成 PackageFileVersion，Inno Setup 7.1.1 編譯成功。預處理輸出確認繁中、英文各含套件及目前已登錄 DLL 的不同標籤，套件檔案版本展開為 1.2.4.2、組件為 1.2.4.0，已登錄檔案版本仍由現有動態查詢讀取。
- 通過：原 COM 公開介面、識別與簽章相對 ISSUE-0005 候選一致；原 API 測試與 RGB 儲存後檢查通過；PowerShell 文字顏色範例以新 DLL 執行成功。以上靜態契約不等同動態 COM 或互動 UI 驗收。
- 未執行：繁中／英文實際 UI、全新及不同已登錄狀態、鍵盤／捲動驗收、實際安裝升級及回復測試。已準備於同一 Sandbox 執行的 `.tools/installer-tests/issue-0005-authorized/results/issue-0006-ui/UiProbe.ps1` 與僅顯示畫面之編譯品，等待使用者協助執行。該測試編譯品並非正式打包產物。
- 未執行：打包。主機先前安裝器 `ExcelKit-Setup-1.2.4-20261116223400-TEST` 仍在執行；已請使用者正常關閉，不自行終止。

測試未完成，未宣稱整體通過或產生正式測試包；保留既有產品及工作區變更。

## 2026-11-17 使用者指定隱藏組件版本
使用者提供第一版繁中畫面截圖，目視可見套件 DLL 檔案版本 1.2.4.2、目前已登錄版本 1.2.4.0；UiProbe v1 因未取得 UI Automation ValuePattern 而失敗，並非顯示測試通過證明。其後使用者明確要求不要顯示組件版本，故前述截圖為舊設計，不可作為新版 UI 驗收。

依新決策已移除套件及 x86/x64 區塊的組件版本與版本比對文字，保留 RegVer 與登錄身分一致性檢查。隔離版 Inno Setup 測試編譯、預處理中英文字與檔案版本檢查通過；確認舊組件版本畫面字串消失。Release DLL、API、RGB 與版本台帳測試維持通過。主機安裝器已正常關閉，最終打包預檢通過；按使用者要求先測試，尚未執行正式測試包打包。

互動 UI 尚未執行：已準備 `.tools/installer-tests/issue-0005-authorized/results/issue-0006-ui/UiProbe2.ps1`，由使用者於現有 Sandbox 執行並取得繁中、英文視窗圖片及檢查結果；尚無結果。實際安裝升級、回復、全新安裝與鍵盤／捲動未執行。電腦操作工具初始化失敗，未在主機執行測試安裝程式。待 UI 測試後始進行打包與報告核對。

## 2026-11-17 修正版 UI 實測與測試工具修正
使用者於 Sandbox 執行 UiProbe2；程式返回 FAIL（Chinese UI text check failed），但測試保存的 `v2-chinesetraditional.txt` 實際包含套件 DLL 檔案版本 1.2.4.2、x86/x64 目前已登錄 DLL 檔案版本 1.2.4.0，無「組件版本」；以獨立 Python 檢查通過並檢視 `v2-chinesetraditional.png`。原因：UiProbe2.ps1 為無 BOM UTF-8，Windows PowerShell 5.1 以本機編碼讀取造成中文正則亂碼，並非畫面顯示錯誤。已建立 BOM 編碼的 UiProbe3 並通過語法解析，保留 v2 失敗證據。英文互動畫面尚未測試，等待使用者在同一 Sandbox 執行 v3；尚未正式打包。

## 2026-11-17 雙語畫面測試與候選包
使用者於同一 Windows 11 x64 Sandbox 執行帶 BOM 的 UiProbe3；`result-v3.txt` 為 PASS、`checks-v3.txt` 記錄繁中與英文均通過。實際視窗文字、截圖均已審查：套件 DLL 檔案版本 1.2.4.2、x86/x64 目前已登錄檔案版本 1.2.4.0；沒有組件版本文字，兩種語言可讀且可捲動。此測試為舊版已登錄狀態；全新未登錄、相同版本、鍵盤操作與其他尺寸仍未驗證。

GUI 測試後以 `Package-ExcelKit.ps1 -Preflight -InstallerVersion 1.2.4+20261117065015` 和相同版本正式流程建立測試安裝包，主機舊安裝器預先確認關閉。產品版本 `1.2.4+20261117065015`、FileVersion `1.2.4.2`，內部 AssemblyVersion `1.2.4.0`。安裝包 SHA-256 `C52FD0012D64901744F1EE4ACEDBF91BAA5FC6E0BD2F60D459C15107AD5DF9D1`，manifest SHA-256 `16A2DE081504D8C017401252A20B9EDBF1279714519F712ABA73E1D8FA03176C`。實際打包 DLL FileVersion、打包 metadata、build report 與 manifest 一致；26 個載荷、30 個來源與報告雜湊一致。實際打包畫面來源與受測 fixture 一致，Inno 預處理確認檔案版本展開 1.2.4.2 且不顯示組件版本。`Installer/Tests/Test-Package-Workflow.ps1` 完整性與負面檢查通過。

實際候選包隔離安裝、x86/x64 COM 與舊版回復測試已準備於 `.tools/installer-tests/issue-0005-authorized/results/package-test-20261117065015/`；等待使用者在同一 Sandbox 執行。候選包已建立，但實際安裝與回復尚未驗證；不以 ISSUE-0005 舊包測試取代本次結果。

## 2026-11-17 實際候選安裝包隔離測試
使用者在已調整測試用原則的 Windows 11 x64 Sandbox 執行 `.tools/installer-tests/issue-0005-authorized/results/package-test-20261117065015/Run-Package-Test.ps1`；result.txt 為 PASS。

- 通過：候選安裝器 SHA-256 與 build report 一致；舊版已安裝 DLL 雜湊確認；新版安裝器升級 exit 0。安裝後 DLL SHA-256 `441FD45854315557C33BDFAAA09FBAF25A7C3D4CFDA08D95BE5CD60CEADE654B` 與套件清單一致，FileVersion 1.2.4.2。
- 通過：安裝新版後 x86／x64 COM 均回報產品版本 1.2.4+20261117065015，既有讀寫及文字顏色方法可呼叫；Verify-ExcelKit.ps1 使用同版 manifest hash 與 Functional，overall pass，12 項 pass。
- 通過：舊測試安裝器回復 exit 0；DLL 雜湊為 `30245AB1D97D3497ED950643AE1FF02BB6248D919CBA954E3EC47A91DA907750`、FileVersion 1.2.4.0；x86／x64 COM 與同版 manifest 加 Functional 驗證 overall pass，12 項 pass。
- 通過：`new-x86.xlsx`、`new-x64.xlsx`、`old-x86.xlsx`、`old-x64.xlsx` 以 openpyxl 重開，A1 文字與紅色、B2 藍色正確。中英文 UI 圖與實際讀取文字、安裝包載荷與來源雜湊均已核對。
- 尚未執行：互動 UI 的全新未登錄與同版已登錄情境、鍵盤焦點與不同尺寸／捲動操作；LabVIEW 實際整合與 32 位元 Windows 亦未執行。以上不得視為通過。

結論：實際安裝包升級、x86／x64 COM、舊版回復及目前舊版已登錄狀態的中英 UI 通過；完整 UI 驗收仍待補測。測試限定 Sandbox，沒有安裝到主機。測試包 SHA-256 `C52FD0012D64901744F1EE4ACEDBF91BAA5FC6E0BD2F60D459C15107AD5DF9D1`。

## 2026-11-17 未登錄與相同版本 UI
使用者在 Sandbox 執行 `issue-0006-states/Run-UiStates.ps1`，result.txt 為 PASS。相同版本狀態下，繁中與英文的套件檔案版本正確、x86/x64 已登錄 DLL 版本與 SHA-256 比對均正確，未顯示組件版本；解除安裝後，兩種語言的 x86/x64 未登錄狀態提示正確。測試最後舊版安裝器回復 exit 0。使用者隨後指示不再檢查、直接打包；保留已建立的 1.2.4+20261117065015 測試包，不重建、不覆寫。鍵盤焦點、不同尺寸與實際捲動操作未執行；不得聲稱通過。

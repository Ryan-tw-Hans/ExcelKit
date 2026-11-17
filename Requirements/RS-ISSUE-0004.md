# ISSUE-0004 儲存格文字顏色需求草案 v0.1

狀態：待人類確認；執行者 Codex（PM）。等級 L3，因新增 COM 跨程式介面。使用者要求 DLL 能修改儲存格文字顏色。

## 建議範圍與驗收
- 新增 `SetCellFontColor(sheetName, cellAddress, red, green, blue)` 及 `SetCellFontColorByRC(sheetName, row, column, red, green, blue)`；RGB 各為 0–255 整數。只設定整個儲存格的字型顏色，不修改值或其他格式；不涵蓋部分字元、背景色、主題色及讀取顏色。
- 無效 RGB、工作表或儲存格位置應回報錯誤；不可將無效 RGB 靜默截斷。儲存並重新開啟 xlsx 後顏色需保留。
- 原有 API 與 DispId 不變，新增唯一 DispId；以 .NET 測試及可用環境下的 COM/VBScript 測試驗證，包括舊方法回歸、邊界值及錯誤。LabVIEW 實測依環境可用性回報，未執行不算通過。
- 更新使用說明與範例；不含部署或安裝。程式資料夾 `src/Excelopen`；既有固定發布資料夾為 `publish/bin/net472` 和 `publish/bin/net10.0`；版本及測試版輸出策略於開發階段依流程確認。

## 五項適用性
架構：適用，新增 COM 公開介面；見 Architecture/AR-ISSUE-0004.md。UI：不適用，無 GUI；字型顏色為資料格式。資料庫：不適用，無業務 DB 變更。系統整合：適用，COM 呼叫端相容性需驗證。維運：不適用，無新增長期服務或維運機制。

## 限制與風險
既有工作區位於 ISSUE-0003 且含未提交變更，開發前需保留與隔離；不得覆蓋既有發布品。部署、現場 LabVIEW 與實際客戶檔案驗證不在已確認範圍。測試計畫與版本策略待後續完成；此草案未經核准。

## 人類決策與實作紀錄
使用者於 2026-11-16 同意需求與架構 v0.1，並要求先測試、更新及執行範例、更新手冊，再打包。分支 feature/ISSUE-0004-font-color；main 起點 1ad27cb8a04e6a66f022e9fda164c999e8fafd26。測試版 1.2.4+20261116213800；結果見 TestReports/TR-ISSUE-0004.md。安裝與 COM 整合受阻，未結案。

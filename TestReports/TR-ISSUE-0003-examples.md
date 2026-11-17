# 手冊與範例安裝測試｜ISSUE-0003

日期：2026-11-15（Asia/Taipei）；執行者 Codex（軟體／測試，單一 Agent 自查）。版本 `1.2.4+20261115094510`，正式版基準仍為 1.2.4。分支 feature/ISSUE-0003-installer，未提交、合併或推送；既有未提交工作保留。

## 本次變更
既有安裝包已包含 PDF、HTML、Markdown 三格式手冊及 VBS／LabVIEW 範例；本次補上開始功能表的手冊與範例資料夾入口、PowerShell 基本 COM 範例及範例使用說明。VBS 範例的開發機 LibreOffice 路徑改為註解形式的一般安裝路徑，同時修正 Err.Description 拼字；其完整 PDF 流程未執行。PowerShell 基本範例預設在使用者 TEMP 下建立 xlsx，不自動列印或匯出 PDF。

安裝程式仍固定在 publish/Installer/Setup，設定與說明在 publish/Installer/Config。安裝後位於所選安裝目錄下 Manual 與 Examples。未改 COM 交易演算法；原交易驗證以 TR-ISSUE-0003-transaction.md 為依據。架構及系統整合沿用既有安裝流程；UI 使用既有雙語標準開始功能表入口；DBA 不適用。

## 驗證結果
- 通過：`dotnet build src/Excelopen.slnx -c Release --no-restore`；兩框架 0 錯誤、23 個既有 net10 可空性警告。ISCC 7.1.1 編譯新版安裝檔成功。
- 通過：Windows 11 x64 Sandbox 靜默安裝（繁中、雙位元）；3 個手冊、5 個範例及說明檔（合計 8 檔）逐一 SHA-256 與發布來源一致。
- 通過：以 32 位元及 64 位元 Windows PowerShell 分別執行安裝後的 QuickStart.ps1，載入本次版本，寫入、讀回儲存格及儲存 xlsx，並核對 xlsx 中存在工作表項目。
- 通過：讀取已安裝 .lnk 的目標，手冊資料夾、範例資料夾及原 PDF 捷徑均指向正確位置。這是捷徑檔案屬性驗證，未做 Explorer 點擊或 PDF 閱讀器視覺驗收。
- 通過：解除安裝成功，已安裝的 PDF 與範例移除，使用者自行建立的 Examples/user-created.txt 保留。
- 未執行：VBS 引擎、LabVIEW、LibreOffice PDF／列印，以及先前仍受操作工具啟動錯誤阻擋的互動精靈視覺驗收。

證據：`.tools/installer-tests/sandbox/results/examples-20261115-094821/`，result.txt=PASS。安裝程式 SHA-256：`A37F7FAD47828A5E55D5A975A2A2CCAD927C8B202D828165B3B0BC90976BAFCE`。測試腳本保存於 Installer/Tests/Examples-Sandbox.ps1。未在主機執行本次新安裝程式。

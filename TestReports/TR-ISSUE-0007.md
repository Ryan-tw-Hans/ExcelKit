# ISSUE-0007 打包版本與影片排除測試報告

2026-11-17，Codex 依序執行 PM、SA、系統整合、維運、開發及測試，自查非獨立審查。基於 main `1ad27cb8a04e6a66f022e9fda164c999e8fafd26`，分支 `feature/ISSUE-0007-package-naming`；工作區含先前未提交內容。使用者明確指定檔名包含 DLL FileVersion、日期與序號，排除影片，修改腳本與 skills，測試後打包。

## 修改與版本
安裝器版本 `1.2.4.3-20261117-01`，檔名 `ExcelKit-Setup-1.2.4.3-20261117-01.exe`，DLL FileVersion `1.2.4.3`，DLL 產品版本 `1.2.4+20261117094415`、AssemblyVersion `1.2.4.0`，正式基準 `1.2.4`。不含 TEST 僅為名稱規則，仍是測試版。PackageVersion.ps1 負責台北日序號及工作區互斥、已用保留紀錄；跨工作區或機器必須協調。指定影片排除打包，原影片及舊安裝包保留。

## 測試
- 通過：`Installer/Tests/Test-PackageVersion.py` 驗證 01／02、預檢唯讀、跨 DLL 同日序號、跨日、舊時間戳檔名共存、部分產物、過期候選、非法日期／序號及兩個並行預留僅一成功。失敗已預留的序號不回收。
- 通過：`Installer/Tests/Test-Package-Preflight.ps1` 測試有效預檢、無副作用、缺少編譯器、版本不符、缺素材及部分產物衝突；獨立 fixture，不改使用者資料。
- 通過：Release 建置 net472/net10.0 0 錯誤，既有 23 項 nullable 警告；原 API、文字顏色測試及 PowerShell 範例通過，輸出活頁簿已重開檢查。
- 通過：Inno Setup 測試編譯並確認 ProductVersion `1.2.4.3-20261117-01`。指定影片來源 SHA-256 `C7EEBCAA896464A628CB3F9F7EABB620F1BCA7E320BBCE8C26D50C896F28F2D7` 前後一致；新包 staging、清單、來源輸入及編譯紀錄均無該影片，其他媒體圖檔保留。25 個載荷、30 個輸入與 build report 雜湊一致。
- 通過：實際 EXE SHA-256 `E2ACC10726CAD9133DE67652582FD0DAF7281E93BD1DCECBA5071B3F9C8369A2`；manifest SHA-256 `083FDAC696800F97F6A9BA6AC4F3265ECC7EB0D00BB33DB9448C8DA71F2C0C95`。打包後再預檢建議 `1.2.4.3-20261117-02`，未建立該版本。`Installer/Tests/Test-Package-Workflow.ps1` 完整性及負面檢查通過。
- 通過：`.agents/skills/excelkit-package`、`.agents/skills/build-release`、`AgentUse/skills/build-release` 的 skill-creator quick_validate。主機 Python 初缺 PyYAML，改用專案既有 `.tools/skill-validation-deps` 及 UTF-8 模式後通過；兩份 build-release 備援內容一致。
- 通過：使用者執行 Sandbox 實際安裝測試。舊時間戳測試包升級到新版本 exit 0；選取手冊及範例後指定影片未安裝，其他媒體安裝；新 DLL FileVersion 1.2.4.3；x86/x64 COM 呼叫與保存活頁簿通過；同版 manifest 驗證器 Functional overall pass（22 pass）。舊測試包回復 exit 0，DLL 回到 FileVersion 1.2.4.0，x86/x64 COM 與 Functional overall pass（12 pass）；四份活頁簿重開後文字及 RGB 正確。證據位於 `.tools/installer-tests/issue-0005-authorized/results/package-test-20261117-01/`。
- 通過：使用者執行 Sandbox 中英文歡迎頁文字測試，修正後 UiProbe2.ps1 的 result-v4.txt 為 PASS，checks-v4.txt 記錄繁體中文與英文皆 PASS；檢查安裝版本、套件及已登錄 DLL 檔案版本與組件版本不顯示。第一次工具測試繁體中文通過，但切換英文時取得無效視窗尺寸而失敗；修正工具等待有效視窗後重測通過。證據位於 `.tools/installer-tests/issue-0005-authorized/results/issue-0007-ui/`。
- 未執行：畫面人工視覺與鍵盤操作、升級前已存在指定影片的情境、LabVIEW、32 位元 Windows 與強制斷電。之前的 ISSUE-0006 UI 測試不能代替本包驗收。

## 限制
測試僅在 Windows 11 x64 Sandbox、兩種位元客戶端及工作區獨立 fixtures 執行；未安裝至主機。升級前已存在該影片的情境尚未實測，不能確認既有檔案是否保留或移除。本次不含正式發布、合併及推送。

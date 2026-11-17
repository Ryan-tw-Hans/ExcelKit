# ISSUE-0004 測試結果

測試者 Codex；軟體、SA 與測試角色由同一執行者依序執行，非獨立審查。2026-11-16（台北）。測試版 1.2.4+20261116213800，正式基準 1.2.4，Assembly/File 1.2.4.0。基礎 main 1ad27cb8a04e6a66f022e9fda164c999e8fafd26，工作區含 ISSUE-0003 等既有未提交內容。

- 通過：`dotnet build src/Excelopen/Excelopen.csproj -c Debug --no-restore -p:PublishToFolder=false`，net472 與 net10.0，0 錯誤；net10 既有 nullable 警告 23 項。打包腳本以 .NET SDK 10.0.403 在獨立 staging 還原及 Release 建置兩個 TFM，Inno Setup 7.1.1 編譯成功。
- 通過：`dotnet run --project TestReports/ISSUE-0004-Checks/Checks.csproj -p:PublishToFolder=false -- <輸出.xlsx> <來源.xlsx> <另存.xlsx> <對照.xlsx>`；原 26 個 DispId 相對 main 均未變，新方法為 27、28；兩種定位、RGB 0/255 和一般值、負數及 256、無效工作表／地址／行列、既有值讀寫通過。以 openpyxl 重開活頁簿，確認 RGB、值與工作表數；與未改顏色的 ClosedXML 另存對照比較，值、字型基本屬性、填色與對齊一致；來源檔保留。
- 注意：ClosedXML 另存會正規化部分預設對齊 XML 屬性；對照組亦同，不能宣稱位元組完全不變。初測缺少工作表時發現既有 GetSheet 自動建立工作表；新方法已修正為回報錯誤，重測通過。測試程式初次比較跨活頁簿 style proxy 造成失敗，改用另存對照檢查後通過。
- 通過：`Reference/PowerShell/FontColor.ps1 -DllPath <net472 DLL> -OutputDirectory <暫存目錄>` 先使用 Debug DLL，再使用本次 staging Release DLL 執行；openpyxl 確認紅色 A1、藍色 B2 與文字。這是 .NET 直接載入，並非 COM 註冊測試。
- 通過：MD、HTML、PDF 已補充說明；PDF 由 HTML 列印，PyMuPDF 產生縮圖，檢查全頁概況及新增內容頁。手冊及範例已在打包前更新。
- 通過：Package-ExcelKit.ps1 預檢及打包；比對安裝器、manifest SHA-256、全部來源輸入與 26 個 staging 清單項目。版本、工具鏈、必要 DLL、範例和手冊在清單中一致。安裝器 SHA-256 E9E22670372D55924EF0CDE68E433C0428E6F1BBFEBCDD19BB18218E765C0C29；manifest SHA-256 13BA75AC7CB7C1BD5C88F74301C0BEA5EAAE6A46747C209CCC3B457C2724EE09。
- 未執行／受阻：網路停用 Windows Sandbox 已啟動，但 Application Control policy 阻擋本版安裝器執行；證據 `.tools/installer-tests/issue-0004/results/result.txt`。故本版安裝、COM x86/x64 實際呼叫、卸載、升降版與回復均未驗證，不得以舊版結果代替。
- 未執行：LabVIEW、互動安裝 UI、32 位元 Windows、現場既有使用者資料。沒有在主機安裝或修改 COM 註冊。效能量測無目標，未執行；資料庫 Migration、UI 改動、持續服務維運不適用。

結果：功能、範例及打包靜態檢查通過；安裝與跨程式驗收受阻，整體尚未結案或合併。測試版位於 `publish/Installer/Setup/ExcelKit-Setup-1.2.4-20261116213800-TEST.exe`，同行保留 build 與 manifest JSON。

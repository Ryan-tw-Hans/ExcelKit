# 可選安裝內容測試結果｜ISSUE-0003

日期 2026-11-16；執行者 Codex（軟體／測試／SA，自查，非獨立審查）。安裝器測試版 `1.2.4+20261116091614`，DLL 組件版本不變，正式版基準 1.2.4。

- **通過**：Inno Setup 7.1.1 以 `publish/Installer/Config/Rebuild-Setup.ps1` 編譯；有既存 `FAULT` 僅在測試編譯使用的未使用變數提示，無編譯錯誤。安裝器 SHA-256 `D9102C6ADF55AEED285F98CF9CDC86C82B4D4FBD4D3E0D2B6955DC6F36AC20CE`，產品版本資訊正確。
- **通過（靜態）**：三任務獨立綁定，手冊與範例初次預設勾選、net10 預設未勾；Media 隨手冊，手冊／範例捷徑各受相應任務控制，net472 為必要項；中英文訊息鍵存在。列舉八種選擇對應檔案規則。維護來源及 Config 的 .iss 雜湊一致；交易助手副本雜湊一致。
- **未執行／受阻**：Windows Sandbox 已啟動並執行 `Installer/Tests/OptionalContent-Sandbox.ps1`，啟動測試安裝檔時系統回報 `An Application Control policy has blocked this file`，故八種組合、COM 實際呼叫、升級、解除安裝及使用者資料保留沒有取得結果。錯誤證據：`.tools/installer-tests/optional/results/optional-20261116091614/result.txt`。未繞過原則。
- **未執行**：雙語互動頁、鍵盤操作、預設值的實際 UI 驗收、既有版本升級、COM 舊版移轉與失敗回復回歸、32 位元 Windows。上版的 Sandbox 結果不能代替本版驗證。維運就緒及合併驗收尚未成立。

本版未在主機執行安裝、註冊或解除安裝。解除安裝既有檔案及重新安裝時捷徑的實際效果亦待隔離測試確認。

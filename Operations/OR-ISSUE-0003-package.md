# 可重複打包維運紀錄｜ISSUE-0003

- 打包入口 `Installer/Package-ExcelKit.ps1`；`Rebuild-Setup.ps1` 轉呼叫相同入口。專案 skill `.agents/skills/excelkit-package/SKILL.md`。
- 每次保留 `.tools/package-work/<version>-<id>/`；結果為固定 Setup 目錄中的 exe、manifest 與 build report。無自動刪除；空間、保留期限與手動清理決策由維護人員依現行規範處理。
- 報告記錄來源、輸出、Inno 與鎖定依賴雜湊，包含版本、工具及測試狀態。缺失的報告或不完整的輸出組應先核對，不可只憑 exe 宣稱打包完整。工具版本或正式基準變更應經正式變更流程，不得為通過預檢逕自改寫。
- 若失敗時已有新檔寫入 Setup，保留既有檔與 staging、檢查報告及版本衝突後人工處理。腳本不終止非本次啟動的程序。
- 驗證腳本預設只讀；需要 ProgramData 交易狀態的讀取權限；拒絕讀取為 unknown，不修改 ACL。外部清單雜湊是完整性對照，並非簽章或身分認證。
- 維運就緒：未確認。需要在可用隔離 Windows 測試實際安裝後記錄、選項組合、重裝／升級、COM 版本與故障回復、功能煙霧及解除安裝。
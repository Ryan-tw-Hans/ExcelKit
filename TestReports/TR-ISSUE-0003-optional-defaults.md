# 可選內容預設值測試結果｜ISSUE-0003

版本 1.2.4+20261116102933，測試版，正式基準 1.2.4。執行者 Codex（PM、軟體、測試及 SA，非獨立審查）。

- 通過：Inno Setup 7.1.1 編譯 publish/Installer/Config/ExcelKit.iss，僅既有 FAULT 未使用提示。安裝檔 SHA-256：3A521DCF711565B4A3D5BEF4AEAB9FD722C1CF7BBC8AC4EF0C82945B50580C15；產品版本資訊一致。
- 通過（靜態）：manual、examples、net10 均有 unchecked；Manual、Media 綁 manual，Examples 綁 examples；UsePreviousTasks=yes；兩個 LibreOffice 完成頁選項保留；維護來源與 Config 副本雜湊一致。
- 未執行：首次安裝中英文 GUI、重新安裝任務狀態、檔案實際部署及移除、靜默安裝、LibreOffice 連結操作、COM 註冊及交易回復。先前上版在 Windows Sandbox 遇到 Application Control 阻擋，不能將該結果當作此版測試結果；本版未在主機執行安裝。
- SA／系統整合／維運：變更僅影響任務預設，未更動檔案綁定；實際安裝、升級及回復仍待隔離測試，維運就緒未確認。資料庫不適用。變更及測試版未提交、未合併。
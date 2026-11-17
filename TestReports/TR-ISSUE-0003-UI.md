# 互動介面測試狀態｜ISSUE-0003

日期：2026-11-15（Asia/Taipei）；執行者：Codex（測試工程師，自查）。受測安裝檔版本 `1.2.4+20261115075927`，SHA-256 `9E11F6BD6BCAC6298A2A139D23D2181C440B97A629CDCA6BA0D00A873B198D03`。

## 本次要求
在 Windows Sandbox 互動確認繁體中文、英文、x86／x64 選擇與舊版移除確認對話框。沿用既有 ISSUE-0003，未擴大程式功能。

## 工具阻塞
依 computer-use skill 初始化 `node_repl` 的 `@oai/sky`；核心程序在初始化時結束，診斷為 `windows sandbox failed: helper_unknown_error: setup refresh had errors`。重設 JavaScript kernel 後再初始化仍相同。尚未取得任何可操作視窗，也沒有執行 UI 點擊、安裝或 COM 變更。

## 結果
- 通過（靜態檢查）：交付設定含繁中及英文，使用 Windows 顯示語言預設、不沿用先前選擇、顯示語言選擇對話框；x86 與 x64 選項設定正確，x64 有 IsWin64 限制，至少選一項；舊版確認使用 Yes/No 且預設 No，包含版本來源資訊；所有實際使用的自訂訊息均有雙語，交付設定與來源相同。
- 未執行：實際語言選擇及切換、版面與文字截斷檢查、x86／x64 勾選與鍵盤操作、空選項提示、舊版確認視窗及按是／否的互動驗收。沒有畫面證據，不宣稱通過。
- 既有 Sandbox 靜默功能測試仍以 `TestReports/TR-ISSUE-0003-transaction.md` 為準，不能代替本次互動驗收。

未修改程式或重新打包；版本與安裝檔雜湊保持一致。下一步需介面操作工具可啟動後，才能繼續本次互動測試。

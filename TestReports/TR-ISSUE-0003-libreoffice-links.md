# LibreOffice 下載連結測試結果｜ISSUE-0003

版本 1.2.4+20261116102441（測試版，正式基準 1.2.4）；執行者 Codex（軟體、測試、SA，自查而非獨立審查）。

- 通過：Inno Setup 7.1.1 編譯 `publish/Installer/Config/ExcelKit.iss`；原有 FAULT 未使用提示，無編譯錯誤。測試安裝檔 SHA-256 `4C5B55DEE23A3304924F0253D371FD5896BA2B615A0F54D9DF064287E0603E4E`；產品版本資訊為 1.2.4+20261116102441。
- 通過（靜態）：兩個網址、shellexec、postinstall、unchecked、skipifsilent、失敗檢查與連結文字檔內容存在；來源及 Config 的 .iss、文字檔與 README SHA-256 一致。完成頁處理先 Finish 交易，失敗時取消勾選；靜默仍在 ssDone 完成交易。此為程式碼檢查，不代表執行已驗證。
- 未執行：中英文完成頁四種選擇、實際瀏覽器開啟、靜默、Apply／Finish 失敗、回滾、升級、解除安裝及 COM 客戶端回歸。先前隔離環境對上版測試安裝檔顯示 Application Control 阻擋（見 TR-ISSUE-0003-optional-content.md），本次尚未取得可執行的隔離環境。本機未執行安裝或修改 COM 註冊。
- 風險：互動安裝的 Finish 交易呼叫提早至按完成時，以便失敗時避免開啟下載頁；其實際順序與失敗回復必須在隔離環境驗證。不能以靜態檢查宣稱完整驗收或維運就緒。
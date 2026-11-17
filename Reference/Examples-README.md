# ExcelKit 安裝後的範例與手冊

從開始功能表的 ExcelKit 項目開啟「操作手冊資料夾」或「範例程式資料夾」。安裝目錄下 `Manual` 保存 PDF、HTML 及 Markdown 操作手冊；`Examples` 依使用工具分類。

- `PowerShell/FontColor.ps1`：以位址及行列號設定 RGB 文字顏色，儲存示例活頁簿。
- `PowerShell/QuickStart.ps1`：基本 COM 範例，建立活頁簿、寫入及讀回儲存格、儲存至目前使用者的 `%TEMP%\ExcelKitExamples`。使用 Windows PowerShell 5.1 執行，不會列印或自動呼叫 LibreOffice。
- `VBS/example.vbs`：包含 PDF 匯出流程的 VBScript 範例。需 Windows 已啟用 VBScript，並依範例確認輸出及日誌資料夾；PDF 匯出另需 LibreOffice。可選的 LibreOffice 路徑已改為一般安裝路徑，請依實際環境設定。
- `LabVIEW/*.vi`：讀取 Excel 與 PDF／列印範例；需相容的 LabVIEW 環境，請先檢查路徑、印表機及 LibreOffice 設定。

PowerShell 範例執行方式（在範例資料夾）：

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PowerShell\QuickStart.ps1

若只註冊 x86 COM，請改用 32 位元 Windows PowerShell：

    C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PowerShell\QuickStart.ps1

範例使用者通常不需管理員權限；安裝與 COM 註冊本身需管理員權限。請選擇與 COM 註冊相符的 32／64 位元執行環境。範例來源放在 Program Files 時，請將要修改的範例複製到自己的工作資料夾。

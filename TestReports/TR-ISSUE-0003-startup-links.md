# 啟動頁與版本偵測測試結果｜ISSUE-0003

受測版本 1.2.4+20261116121333，測試版，正式基準 1.2.4；Codex 擔任軟體、SA、系統整合與測試，均為自查，非獨立審查。

- 通過：Inno Setup 7.1.1 透過 publish/Installer/Config/Rebuild-Setup.ps1 編譯。唯一提示為原有 FAULT 僅測試編譯使用時未使用。產品版本資訊符合；安裝檔 SHA-256 `CB9922318C24061058B47B97B54BA9E8534A054DB2B59ABD44CF60A3D78AD75C`。
- 通過：打包 DLL 的組件版本 1.2.4.0、SHA-256 `54376B9A5A2AC2DF6877E40E8F9102A4D8CDCCA5EF375038A455CD9C44483166` 與產生的 PackageMetadata.iss 一致。獨立測試目錄將雜湊改為零後編譯，預處理器以 exit 2 及預期訊息拒絕，沒有改動來源 DLL。
- 通過（靜態）：Framework 檢查位於 PrepareToInstall 且先於暫存助手抽取、COM Inspect 與 Begin；雙視圖與使用者層註冊檢查、連結使用原使用者執行、三個完成頁選項的預設與靜默條件存在；來源及固定 Config 副本一致。
- 通過（唯讀探測）：本機 x86／x64 機器層 COM 版本子機碼均為 1.2.4.0，登錄指向 DLL 存在；DLL 與打包來源 SHA-256 相同。另以 Windows PathCreateFromUrlW 及登錄相同的 URL 探測，成功轉成既有本機路徑。此探測不等同安裝器內 Pascal 呼叫的執行測試。
- 未執行：啟動頁實際畫面、字體縮放與鍵盤操作，兩個網址與安裝資料夾實際開啟；空白、單／雙位元、不一致、HKCU 覆寫、拒絕讀取與 DLL 缺失的隔離場景；缺 Framework 的 GUI／靜默阻擋、成功安裝、升級、解除安裝、COM 遷移與回復。圖形介面檢查工具本次未能啟動；上一版隔離環境曾有 Application Control 阻擋，但不能推定本版執行結果。
- 風險：實際 UI 高 DPI 排版、Pascal 對 Windows URL 轉換 API 的呼叫、不同 UAC 帳號下 HKCU 視圖、Framework 阻擋與原 COM 交易保護皆待隔離實測。需要此證據才能宣稱驗收通過、維運就緒或申請合併。本次沒有執行安裝、變更 COM 註冊、提交或合併。
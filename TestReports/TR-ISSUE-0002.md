# 測試結果｜ISSUE-0002

執行者：Codex（測試工程師，自查，非獨立審查）；2026-11-14；測試版 1.2.4+20261114220803。

- 通過：發布前 publish 不存在，未發現可辨識之 ExcelKit 執行檔程序；本產品是 DLL，無法只以程序名稱排除其他宿主載入。首次目標不存在，因此不覆寫舊發布 DLL。
- 通過：`dotnet build src/Excelopen.slnx -c Release`，net472 和 net10.0 建置成功，0 錯誤、23 項 net10 nullability 警告。`--no-restore -v quiet` 重複建置成功，增量建置 0 錯誤。
- 通過：Manual 3、VBS 1、LabVIEW 2、Media 2 檔案與 Reference 原始檔逐一 SHA-256 一致。
- 通過：兩 DLL FileVersion 均為 1.2.4.0，ProductVersion 含 1.2.4+20261114220803 與來源 commit；net472 載入後建立活頁簿、寫入及讀回 A1，GetVersionInfo 回傳相同測試版。
- DLL SHA-256：net472 `FB76C857F12C16C233778848D762583B3AAB34CB9997F742DDE35B8629FE0F56`；net10.0 `9DAEF890A7551694AFA5CB8584FC98A1C7BE36B0608376CF393F1F844CD89B19`。
- 未執行：COM 註冊、VBS／LabVIEW 客戶端、net10 執行期及 LibreOffice PDF／列印端對端驗證；未宣稱這些功能通過。
- Migration／回復：應用資料庫 migration 不適用；無舊 publish 須還原，未執行回復演練。

# 需求追溯矩陣（RTM）｜ExcelKit v1.2.4

| 項目 | 內容 |
| --- | --- |
| 文件編號 | RTM-ExcelKit-v1.2.4 |
| 文件版本 | v0.1（草稿） |
| 日期 | 2026-11-14（Asia/Taipei） |
| 撰寫者 | Claude（測試工程師角色；自查，非獨立審查） |
| 依據 | SRS-ExcelKit-v1.2.4、TCS-ExcelKit-v1.2.4、CR-ExcelKit-v1.2.4 |

既有覆蓋判定（僅針對現有 `測試腳本/`）：
- **完整**：既有腳本已驗證主要行為。
- **部分**：既有腳本有呼叫，但驗證不足（例如只檢查無錯誤）。
- **缺口**：無既有腳本。
- **審查**：COM 不可呼叫，以程式審查覆蓋。

執行狀態：所有案例目前皆為「未執行」。

## 1. 功能需求追溯

| 需求 | 說明 | 測試案例 | 既有腳本 | 既有覆蓋 | 相關問題 |
| --- | --- | --- | --- | --- | --- |
| FR-VER-01 | GetVersionInfo | TC-VER-01 | 02 | 部分 | ISS-17 |
| FR-WB-01 | CreateNewWorkbook | TC-WB-01、TC-WB-08、TC-WB-09 | 03 | 部分 | ISS-01、ISS-04、ISS-05 |
| FR-WB-02 | OpenWorkbook | TC-WB-03、TC-WB-06、TC-WB-07、TC-WB-11 | 07、08、11 | 部分（無負面） | ISS-04 |
| FR-WB-03 | SaveWorkbook | TC-WB-02、TC-WB-04、TC-WB-05、TC-WB-09、TC-WB-10、TC-WB-11 | 03、08 | 部分（未測 `""` 覆寫原檔） | ISS-05 |
| FR-SH-01 | AddSheet | TC-SH-01、TC-SH-02、TC-SH-03 | 04 | 部分 | — |
| FR-SH-02 | DeleteSheet | TC-SH-04、TC-SH-05、TC-SH-08 | 12 | 部分 | ISS-08 |
| FR-SH-03 | GetAllSheetNames | TC-SH-09、TC-WB-01 | 19 | 部分 | — |
| FR-SH-04 | 工作表自動建立 | TC-SH-06、TC-SH-07 | — | 缺口 | ISS-03 |
| FR-CL-01 | SetCellValue（A1） | TC-CL-01、TC-CL-06、TC-CL-09、TC-CL-11 | 05 | 部分 | — |
| FR-CL-02 | SetCellValueByRC | TC-CL-02、TC-CL-09 | 06 | 部分 | — |
| FR-CL-03 | GetCellValue（A1） | TC-CL-03、TC-CL-05、TC-CL-10、TC-CL-12 | 07 | 部分（僅字串） | — |
| FR-CL-04 | GetCellValueByRC | TC-CL-04 | 14、15、17（間接） | 部分 | — |
| FR-CL-05 | ClearCell | TC-CL-07 | 13 | 部分（未測格式） | ISS-10 |
| FR-CL-06 | 型別對應 | TC-CL-05、TC-CL-08 | — | 缺口 | ISS-02 |
| FR-CL-07 | FormatCell（.NET） | — | — | 審查 | — |
| FR-RC-01 | InsertRow | TC-RC-01、TC-RC-07、TC-RC-08 | 14 | 完整（正向） | — |
| FR-RC-02 | DeleteRow | TC-RC-02、TC-RC-07 | 15 | 部分（未驗上移） | — |
| FR-RC-03 | Set/GetRowHeight | TC-RC-03、TC-RC-07 | 16 | 部分（未驗持久化） | — |
| FR-RC-04 | InsertColumn | TC-RC-04、TC-RC-07 | 17 | 完整（正向） | — |
| FR-RC-05 | DeleteColumn | TC-RC-05、TC-RC-07 | 17 | 完整（正向） | — |
| FR-RC-06 | Set/GetColumnWidth | TC-RC-06、TC-RC-07 | 18 | 部分（未驗持久化） | — |
| FR-OUT-01 | ExportToPdf | TC-OUT-01～06、TC-OUT-08～10 | 09 | 部分 | ISS-06、ISS-09、ISS-12、ISS-14、ISS-15 |
| FR-OUT-02 | PrintWorkbook | TC-OUT-11、TC-OUT-12、TC-OUT-13 | 10 | 部分（腳本有誤 T-02） | ISS-12 |
| FR-OUT-03 | LibreOffice 路徑解析 | TC-OUT-02、TC-OUT-14 | 09、20 | 部分 | — |
| FR-OUT-04 | 匯出／列印前自動存檔 | TC-OUT-05、TC-OUT-15 | — | 缺口 | ISS-06 |
| FR-OUT-05 | 逾時 30 秒 | TC-OUT-07 | — | 缺口 | — |
| FR-OUT-06 | 三參數 ExportToPdf（.NET） | — | — | 審查 | ISS-11 |
| FR-CFG-01 | SetLibreOfficePath | TC-CFG-01、TC-CFG-08、TC-OUT-14 | 20、09 | 部分 | — |
| FR-CFG-02 | SetLogDirectory | TC-CFG-02、TC-CFG-05 | 20 | 部分（未驗檔案） | ISS-18 |
| FR-CFG-03 | SetLogEnabled | TC-CFG-03 | 20 | 部分 | — |
| FR-CFG-04 | SetLogLevel | TC-CFG-04、TC-CFG-07 | 20 | 部分 | — |
| FR-CFG-05 | 日誌格式與條件 | TC-CFG-02、TC-CFG-04 | — | 缺口 | — |
| FR-RES-01 | Dispose（.NET） | TC-PERF-02（間接） | — | 審查 | ISS-04 |

## 2. 非功能需求與介面追溯

| 需求 | 說明 | 測試案例 | 既有覆蓋 | 相關問題 |
| --- | --- | --- | --- | --- |
| NFR-01 | .NET 4.7.2／x86、x64 | TC-BLD-01、TC-ENV-01、TC-ENV-02、TC-ENV-04 | 部分（僅 x64 預設） | — |
| NFR-02 | net10.0 非 COM | TC-BLD-01 | 缺口 | ISS-16 |
| NFR-03 | 單一 DLL 部署 | TC-BLD-02 | 缺口 | — |
| NFR-04 | 效能 | TC-OUT-07、TC-PERF-01、TC-PERF-02 | 缺口 | ISS-04 |
| NFR-05 | 外部行程逾時終止 | TC-OUT-07 | 缺口 | ISS-12 |
| NFR-06 | 可診斷性 | TC-CFG-02～05 | 部分 | ISS-18 |
| NFR-07 | 執行緒 | （文件約束） | 不適用 | ISS-13 |
| NFR-08 | 覆寫風險 | TC-OUT-05、TC-OUT-06 | 缺口 | ISS-06 |
| NFR-09 | 版本同步 | TC-VER-02 | 缺口 | ISS-17 |
| NFR-10 | 授權 | （發布檢查） | 不適用 | — |
| 4.1 | COM 介面／DispId | TC-BLD-03、TC-ENV-03 | 缺口 | ISS-07 |
| 2.4-2 | 設定行程共用 | TC-CFG-06 | 缺口 | ISS-13 |
| 2.4-4 | 不支援 .xls | TC-WB-07 | 缺口 | — |
| 2.2 | 使用情境（LabVIEW） | TC-INT-01、TC-INT-02、TC-INT-04 | 部分（手動） | — |

## 3. 問題清單 → 驗證案例

| 問題 | 嚴重度 | 驗證案例 |
| --- | --- | --- |
| ISS-01 | 中 | TC-WB-08 |
| ISS-02 | 高 | TC-CL-08 |
| ISS-03 | 中 | TC-SH-06、TC-SH-07 |
| ISS-04 | 低 | TC-PERF-02 |
| ISS-05 | 高 | TC-WB-09 |
| ISS-06 | 高 | TC-OUT-05、TC-OUT-15 |
| ISS-07 | 高 | TC-ENV-03 |
| ISS-08 | 低 | TC-SH-05 |
| ISS-09 | 中 | TC-OUT-08 |
| ISS-10 | 低 | TC-CL-07 |
| ISS-11 | 低 | 程式審查 |
| ISS-12 | 中 | 程式審查（TC-OUT-07 間接） |
| ISS-13 | 中 | TC-CFG-06 |
| ISS-14 | 低 | TC-OUT-03 |
| ISS-15 | 中 | TC-OUT-10 |
| ISS-16 | 低 | TC-BLD-01 |
| ISS-17 | 低 | TC-VER-02 |
| ISS-18 | 低 | TC-CFG-05 |
| ISS-19 | 中 | 腳本整理後重跑附錄 A 所列案例 |
| ISS-20 | 低 | 文件審查 |

## 4. 覆蓋摘要（既有腳本）

| 判定 | 功能需求數（共 33） |
| --- | --- |
| 完整 | 3 |
| 部分 | 22 |
| 缺口 | 5 |
| 審查 | 3 |

主要缺口：型別對應（FR-CL-06）、工作表自動建立（FR-SH-04）、匯出前自動存檔（FR-OUT-04）、逾時（FR-OUT-05）、日誌內容（FR-CFG-05），以及所有負面／邊界案例。三項高嚴重度資料問題（ISS-02、05、06）目前皆無既有測試可偵測。

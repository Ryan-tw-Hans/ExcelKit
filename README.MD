# ExcelKit

`ExcelKit` 是一個以 C# 開發的 Excel 操作元件，提供：

- 以 COM 方式給 LabVIEW / VBScript / 其他 COM Client 呼叫
- 建立、開啟、儲存活頁簿
- Sheet / 儲存格 / 列與欄操作
- 透過 LibreOffice 進行 PDF 匯出與列印
- 針對自動化測試提供完整 VBS 測試腳本

## 專案結構

- `src/Excelopen/`：核心函式庫（`ExcelKit`，多目標 `net472;net10.0`）
- `Reference/`：手冊、VBS 範例、LabVIEW 範例、影片與圖片
- `測試腳本/`：VBScript 測試案例與測試報告輸出
- `src/Excelopen.slnx`：方案檔

## 主要功能

`ExcelOperator`（ProgID: `ExcelKit.ExcelOperator`）提供以下能力：

- 活頁簿操作：`CreateNewWorkbook`、`OpenWorkbook`、`SaveWorkbook`
- Sheet 操作：`AddSheet`、`DeleteSheet`、`GetAllSheetNames`
- 儲存格操作：
  - 依位址：`SetCellValue`、`GetCellValue`、`ClearCell`
  - 依行列：`SetCellValueByRC`、`GetCellValueByRC`
- 列/欄操作：`InsertRow`、`DeleteRow`、`SetRowHeight`、`GetRowHeight`、`InsertColumn`、`DeleteColumn`、`SetColumnWidth`、`GetColumnWidth`
- 輸出與列印：`ExportToPdf`、`PrintWorkbook`
- LibreOffice 設定與日誌：`SetLibreOfficePath`、`SetLogDirectory`、`SetLogEnabled`、`SetLogLevel`

## 建置需求

- Visual Studio 2022/2026 或 .NET SDK
- 目標框架：`.NET Framework 4.7.2`、`.NET 10`
- 若使用 PDF 匯出/列印：需安裝 LibreOffice（`soffice.exe`）

## 快速開始

1. 建置專案
   - 開啟 `src/Excelopen.slnx` 後建置，或在 `src/` 下使用 `dotnet build`。
2. 若需 COM 用途（`net472`）
   - 參考 `Reference/Manual/ExcelKit_User_Manual.md` 中的 COM 註冊指令進行註冊。
3. 執行測試腳本
   - 到 `測試腳本/` 執行 `Run_All_Ordered_Tests.vbs`。

## 參考資料

所有對外參考素材已集中於 `Reference/`，可依類型直接查閱：

- `Reference/Manual/ExcelKit_User_Manual.md`：Markdown 使用手冊
- `Reference/Manual/ExcelKit_User_Manual.html`：HTML 使用手冊
- `Reference/Manual/ExcelKit_User_Manual.pdf`：PDF 使用手冊
- `Reference/VBS/example.vbs`：VBScript 基本範例
- `Reference/LabVIEW/ExcelKit_ReadExistFile.vi`：LabVIEW 讀取既有 Excel 範例
- `Reference/LabVIEW/ExcelKit_SavePDFPrint.vi`：LabVIEW 匯出 PDF 與列印範例
- `Reference/Media/Snipaste_2026-04-03_19-25-28.png`：操作畫面截圖

若需測試完整 API 流程，仍可參考 `測試腳本/` 內的編號 VBS 測試案例。

## 授權

本專案採用 `MIT License`，詳見 `LICENSE`。

---

## 第三方套件與版權使用聲明

本專案使用下列第三方套件（NuGet）。其著作權與商標均屬原作者或權利人所有；
本專案僅依各套件授權條款於原授權範圍內使用。

| 套件 | 版本 | 主要用途 | 授權 |
|---|---:|---|---|
| `ClosedXML` | `0.105.0` | Excel (`.xlsx`) 讀寫與操作 | MIT |
| `Fody` | `6.8.1` | 編譯期 IL weaving 基礎框架 | MIT |
| `Costura.Fody` | `5.7.0` | 相依 DLL 嵌入/合併（特定目標使用） | MIT |
| `System.Diagnostics.DiagnosticSource` | `8.0.0` | 診斷與追蹤 API | MIT (.NET Foundation) |
| `System.Text.RegularExpressions` | `4.3.1` | 正規表示式支援（`net472`） | MIT (.NET Foundation) |
| `System.Net.Http` | `4.3.4` | HTTP 類別庫（`net472`） | MIT (.NET Foundation) |
| `Microsoft.NETCore.Platforms` | `1.1.1` | 平台識別與相容性資訊（`net10.0`） | MIT (.NET Foundation) |

### 補充說明

- 套件授權資訊以 NuGet 套件頁面與其原始碼儲存庫（LICENSE 檔）為準。
- 若您重新散布本專案二進位或原始碼，請一併遵循上述第三方套件授權要求。
- `LibreOffice` 為外部工具，不屬於本專案內建 NuGet 相依；若隨產品散布，請自行遵循 LibreOffice 官方授權條款。
## 發布資料夾

執行 `dotnet build src/Excelopen.slnx -c Release` 會把元件整理到 `publish/bin/net472/` 和 `publish/bin/net10.0/`，手冊放在 `publish/Manual/`，VBS 與 LabVIEW 範例放在 `publish/Examples/`，圖片及影片放在 `publish/Media/`。建置不會自動刪除舊檔；升級前請檢查目標內容，避免保留不再需要的舊版相依檔。COM 用途請參考手冊，建置不會自動註冊元件。

using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

 namespace Excelopen
{
    /// <summary>
    /// COM 接口 - 用于 LabVIEW 和其他 COM 客户端
    /// </summary>
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789012")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExcelOperator
    {
        [DispId(1)]
        void CreateNewWorkbook();
        [DispId(2)]
        void OpenWorkbook(string filePath);
        [DispId(3)]
        void SaveWorkbook(string filePath);
        [DispId(4)]
        void SetCellValue(string sheetName, string cellAddress, object value);
        [DispId(5)]
        object GetCellValue(string sheetName, string cellAddress);
        [DispId(6)]
        void DeleteSheet(string sheetName);
        [DispId(7)]
        string GetVersionInfo();
        [DispId(8)]
        void ExportToPdf(string excelFilePath, string pdfOutputPath);
        [DispId(9)]
        void PrintWorkbook(string excelFilePath, string printerName);
        [DispId(10)]
        void SetLibreOfficePath(string libreOfficePath);
        [DispId(11)]
        void SetLogDirectory(string logDirectoryPath);
        [DispId(12)]
        void SetLogEnabled(bool enabled);
        [DispId(13)]
        void SetLogLevel(int logLevel);
        [DispId(14)]
        void AddSheet(string sheetName);
        [DispId(15)]
        void SetCellValueByRC(string sheetName, int row, int column, object value);
        [DispId(16)]
        object GetCellValueByRC(string sheetName, int row, int column);
        [DispId(17)]
        void ClearCell(string sheetName, string cellAddress);
        [DispId(18)]
        void InsertRow(string sheetName, int rowNumber);
        [DispId(19)]
        void DeleteRow(string sheetName, int rowNumber);
        [DispId(20)]
        void SetRowHeight(string sheetName, int rowNumber, double height);
        [DispId(21)]
        double GetRowHeight(string sheetName, int rowNumber);
        [DispId(22)]
        void InsertColumn(string sheetName, int columnNumber);
        [DispId(23)]
        void DeleteColumn(string sheetName, int columnNumber);
        [DispId(24)]
        void SetColumnWidth(string sheetName, int columnNumber, double width);
        [DispId(25)]
        double GetColumnWidth(string sheetName, int columnNumber);
        [DispId(26)]
        object GetAllSheetNames();
    }

    /// <summary>
    /// Excel 操控類庫，支援 xlsx 檔案的完整操作
    /// COM 注册类 - 用于 LabVIEW 和其他 COM 客户端
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("87654321-4321-4321-4321-210987654321")]
    [ProgId("ExcelKit.ExcelOperator")]
    [ComDefaultInterface(typeof(IExcelOperator))]
    public class ExcelOperator : IExcelOperator, IDisposable
    {
        public ExcelOperator()
        {
        }

        /// <summary>
        /// 庫版本號
        /// </summary>
        public static readonly string Version = "1.2.2";

        /// <summary>
        /// 庫名稱
        /// </summary>
        public static readonly string LibraryName = "ExcelKit";

        /// <summary>
        /// 最後更新日期
        /// </summary>
        public static readonly string LastUpdateDate = "2026-04-05";

        /// <summary>
        /// 取得版本資訊字串 (靜態版本)
        /// </summary>
        public static string GetVersionInfoStatic()
        {
            return $"{LibraryName} v{Version} ({LastUpdateDate})";
        }

        /// <summary>
        /// 取得版本資訊字串 (COM 相容 - 實例方法)
        /// </summary>
        public string GetVersionInfo()
        {
            return GetVersionInfoStatic();
        }

        private XLWorkbook _workbook;
        private string _filePath;

        public const int LogLevelError = 1;
        public const int LogLevelWarning = 2;
        public const int LogLevelInfo = 3;

        /// <summary>
        /// 建立新的 Excel 檔案
        /// </summary>
        public void CreateNewWorkbook()
        {
            _workbook = new XLWorkbook();
            // 添加默认工作表，ClosedXML 要求至少有一个工作表
            _workbook.Worksheets.Add("Sheet1");
        }

        /// <summary>
        /// 開啟現有 Excel 檔案
        /// </summary>
        public void OpenWorkbook(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"檔案不存在: {filePath}");

            _workbook = new XLWorkbook(filePath);
            _filePath = filePath;
        }

        /// <summary>
        /// 儲存 Excel 檔案
        /// </summary>
        public void SaveWorkbook(string filePath = null)
        {
            if (string.IsNullOrEmpty(filePath))
                filePath = _filePath;

            if (string.IsNullOrEmpty(filePath))
                throw new InvalidOperationException("必須提供檔案路徑");

            _workbook.SaveAs(filePath);
            _filePath = filePath;
        }

        #region Sheet 操作

        /// <summary>
        /// 新增 Sheet (COM 相容 - 隱含實作介面)
        /// </summary>
        void IExcelOperator.AddSheet(string sheetName)
        {
            AddSheet(sheetName);
        }

        /// <summary>
        /// 新增 Sheet
        /// </summary>
        public IXLWorksheet AddSheet(string sheetName)
        {
            return _workbook.Worksheets.Add(sheetName);
        }

        /// <summary>
        /// 刪除 Sheet
        /// </summary>
        public void DeleteSheet(string sheetName)
        {
            _workbook.Worksheets.Delete(sheetName);
        }

        /// <summary>
        /// 取得 Sheet (如果不存在则自动创建)
        /// </summary>
        public IXLWorksheet GetSheet(string sheetName)
        {
            try
            {
                return _workbook.Worksheets.Worksheet(sheetName);
            }
            catch
            {
                // 如果工作表不存在，自动创建
                return _workbook.Worksheets.Add(sheetName);
            }
        }

        /// <summary>
        /// 取得所有 Sheet 名稱 (COM 相容)
        /// </summary>
        object IExcelOperator.GetAllSheetNames()
        {
            return GetAllSheetNames().ToArray();
        }

        /// <summary>
        /// 取得所有 Sheet 名稱
        /// </summary>
        public List<string> GetAllSheetNames()
        {
            return _workbook.Worksheets.Select(w => w.Name).ToList();
        }

        #endregion

        #region 儲存格操作

        /// <summary>
        /// 寫入儲存格值（自動支援常見資料類型）
        /// </summary>
        public void SetCellValue(string sheetName, string cellAddress, object value)
        {
            var sheet = GetSheet(sheetName);

            if (value == null)
            {
                sheet.Cell(cellAddress).Clear();
            }
            else if (value is string strValue)
            {
                sheet.Cell(cellAddress).Value = strValue;
            }
            else if (value is int intValue)
            {
                sheet.Cell(cellAddress).Value = intValue;
            }
            else if (value is double doubleValue)
            {
                sheet.Cell(cellAddress).Value = doubleValue;
            }
            else if (value is decimal decimalValue)
            {
                sheet.Cell(cellAddress).Value = (double)decimalValue;
            }
            else if (value is bool boolValue)
            {
                sheet.Cell(cellAddress).Value = boolValue;
            }
            else if (value is DateTime dateValue)
            {
                sheet.Cell(cellAddress).Value = dateValue;
            }
            else if (value is long longValue)
            {
                sheet.Cell(cellAddress).Value = longValue;
            }
            else if (value is float floatValue)
            {
                sheet.Cell(cellAddress).Value = floatValue;
            }
            else
            {
                // 對於未知類型，嘗試轉換為字符串
                sheet.Cell(cellAddress).Value = value.ToString();
            }
        }

        /// <summary>
        /// 讀取儲存格值
        /// </summary>
        public object GetCellValue(string sheetName, string cellAddress)
        {
            var sheet = GetSheet(sheetName);
            var value = sheet.Cell(cellAddress).Value;
            return ConvertXLCellValue(value);
        }

        private object ConvertXLCellValue(XLCellValue value)
        {
            if (value.IsBlank) return null;
            if (value.IsBoolean) return value.GetBoolean();
            if (value.IsNumber) return value.GetNumber();
            if (value.IsDateTime) return value.GetDateTime();
            if (value.IsError) return value.GetError().ToString();
            return value.ToString();
        }

        /// <summary>
        /// 寫入儲存格值（按行列號，自動支援常見資料類型）
        /// </summary>
        public void SetCellValueByRC(string sheetName, int row, int column, object value)
        {
            SetCellValue(sheetName, row, column, value);
        }

        /// <summary>
        /// 寫入儲存格值（按行列號）
        /// </summary>
        public void SetCellValue(string sheetName, int row, int column, object value)
        {
            var sheet = GetSheet(sheetName);

            if (value == null)
            {
                sheet.Cell(row, column).Clear();
            }
            else if (value is string strValue)
            {
                sheet.Cell(row, column).Value = strValue;
            }
            else if (value is int intValue)
            {
                sheet.Cell(row, column).Value = intValue;
            }
            else if (value is double doubleValue)
            {
                sheet.Cell(row, column).Value = doubleValue;
            }
            else if (value is decimal decimalValue)
            {
                sheet.Cell(row, column).Value = (double)decimalValue;
            }
            else if (value is bool boolValue)
            {
                sheet.Cell(row, column).Value = boolValue;
            }
            else if (value is DateTime dateValue)
            {
                sheet.Cell(row, column).Value = dateValue;
            }
            else if (value is long longValue)
            {
                sheet.Cell(row, column).Value = longValue;
            }
            else if (value is float floatValue)
            {
                sheet.Cell(row, column).Value = floatValue;
            }
            else
            {
                // 對於未知類型，嘗試轉換為字符串
                sheet.Cell(row, column).Value = value.ToString();
            }
        }

        /// <summary>
        /// 讀取儲存格值（按行列號）
        /// </summary>
        public object GetCellValueByRC(string sheetName, int row, int column)
        {
            return GetCellValue(sheetName, row, column);
        }

        /// <summary>
        /// 讀取儲存格值（按行列號）
        /// </summary>
        public object GetCellValue(string sheetName, int row, int column)
        {
            var sheet = GetSheet(sheetName);
            var value = sheet.Cell(row, column).Value;
            return ConvertXLCellValue(value);
        }

        /// <summary>
        /// 刪除儲存格內容
        /// </summary>
        public void ClearCell(string sheetName, string cellAddress)
        {
            var sheet = GetSheet(sheetName);
            sheet.Cell(cellAddress).Clear();
        }

        /// <summary>
        /// 設定儲存格格式（如字型、背景色等）
        /// </summary>
        public void FormatCell(string sheetName, string cellAddress, Action<IXLStyle> styleAction)
        {
            var sheet = GetSheet(sheetName);
            styleAction?.Invoke(sheet.Cell(cellAddress).Style);
        }

        #endregion

        #region 行操作

        /// <summary>
        /// 新增行
        /// </summary>
        public void InsertRow(string sheetName, int rowNumber)
        {
            var sheet = GetSheet(sheetName);
            sheet.Row(rowNumber).InsertRowsAbove(1);
        }

        /// <summary>
        /// 刪除行
        /// </summary>
        public void DeleteRow(string sheetName, int rowNumber)
        {
            var sheet = GetSheet(sheetName);
            sheet.Row(rowNumber).Delete();
        }

        /// <summary>
        /// 設定行高
        /// </summary>
        public void SetRowHeight(string sheetName, int rowNumber, double height)
        {
            var sheet = GetSheet(sheetName);
            sheet.Row(rowNumber).Height = height;
        }

        /// <summary>
        /// 取得行高
        /// </summary>
        public double GetRowHeight(string sheetName, int rowNumber)
        {
            var sheet = GetSheet(sheetName);
            return sheet.Row(rowNumber).Height;
        }

        #endregion

        #region 列操作

        /// <summary>
        /// 新增列
        /// </summary>
        public void InsertColumn(string sheetName, int columnNumber)
        {
            var sheet = GetSheet(sheetName);
            sheet.Column(columnNumber).InsertColumnsBefore(1);
        }

        /// <summary>
        /// 刪除列
        /// </summary>
        public void DeleteColumn(string sheetName, int columnNumber)
        {
            var sheet = GetSheet(sheetName);
            sheet.Column(columnNumber).Delete();
        }

        /// <summary>
        /// 設定列寬
        /// </summary>
        public void SetColumnWidth(string sheetName, int columnNumber, double width)
        {
            var sheet = GetSheet(sheetName);
            sheet.Column(columnNumber).Width = width;
        }

        /// <summary>
        /// 取得列寬
        /// </summary>
        public double GetColumnWidth(string sheetName, int columnNumber)
        {
            var sheet = GetSheet(sheetName);
            return sheet.Column(columnNumber).Width;
        }

        #endregion

        #region PDF 和列印操作（透過 LibreOffice）

        /// <summary>
        /// 取得 LibreOffice 配置管理器
        /// </summary>
        private static LibreOfficeConfigManager _configManager;

        /// <summary>
        /// 初始化 LibreOffice 配置（在首次使用時自動呼叫）
        /// </summary>
        private static LibreOfficeConfigManager GetConfigManager()
        {
            if (_configManager == null)
            {
                _configManager = new LibreOfficeConfigManager();
            }
            return _configManager;
        }

        /// <summary>
        /// 匯出為 PDF（使用 LibreOffice - 自動讀取配置）
        /// </summary>
        public void ExportToPdf(string excelFilePath, string pdfOutputPath)
        {
            try
            {
                // 自動同步：如果記憶體中有開啟且修改過的工作簿，先儲存以供 LibreOffice 讀取
                if (_workbook != null)
                {
                    string targetSavePath = string.IsNullOrWhiteSpace(excelFilePath) ? _filePath : excelFilePath;
                    if (!string.IsNullOrEmpty(targetSavePath))
                    {
                        SaveWorkbook(targetSavePath);
                        excelFilePath = targetSavePath;
                    }
                }

                if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
                    throw new FileNotFoundException($"找不到 Excel 檔案: {excelFilePath}");

                // 確保為絕對路徑，避免 LibreOffice 參數解析錯誤
                excelFilePath = Path.GetFullPath(excelFilePath);
                pdfOutputPath = Path.GetFullPath(pdfOutputPath);

                // 取得配置管理器並檢查配置
                var configManager = GetConfigManager();
                configManager.LogOperation($"開始 PDF 匯出：{excelFilePath}", LogLevelInfo);

                if (!configManager.IsEnabled())
                {
                    throw new Exception("LibreOffice 轉換功能已被停用");
                }

                // 取得 LibreOffice 路徑
                string libreOfficePath = configManager.GetLibreOfficePath();

                if (string.IsNullOrWhiteSpace(libreOfficePath))
                {
                    throw new Exception(
                        "LibreOffice 路徑未設定或不存在。\n" +
                        "請先呼叫 SetLibreOfficePath() 指定 soffice.exe。"
                    );
                }

                // 取得超時時間
                int timeout = configManager.GetTimeoutSeconds();

                string outputDir = Path.GetDirectoryName(pdfOutputPath);
                if (string.IsNullOrEmpty(outputDir)) outputDir = Directory.GetCurrentDirectory();

                // 若是磁碟根目錄 (例如 D:\)，不可 TrimEnd 成 D:，改為 D:\.
                // 這也可避免命令列引號在尾端反斜槓時被錯誤跳脫。
                string rootDir = Path.GetPathRoot(outputDir);
                if (!string.IsNullOrEmpty(rootDir) &&
                    string.Equals(Path.GetFullPath(outputDir), Path.GetFullPath(rootDir), StringComparison.OrdinalIgnoreCase))
                {
                    outputDir = Path.Combine(outputDir, ".");
                }
                else if (outputDir.EndsWith("\\"))
                {
                    outputDir = outputDir.TrimEnd('\\');
                }

                // 執行轉換
                var processInfo = new ProcessStartInfo
                {
                    FileName = libreOfficePath,
                    Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{excelFilePath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processInfo))
                {
                    bool completed = process.WaitForExit(timeout * 1000);

                    if (!completed)
                    {
                        try { process.Kill(); } catch { }
                        throw new TimeoutException($"LibreOffice 轉換超時（{timeout}秒）");
                    }

                    if (process.ExitCode != 0)
                    {
                        string errorOutput = process.StandardError.ReadToEnd();
                        throw new Exception($"LibreOffice 轉換失敗 (退出碼 {process.ExitCode})：{errorOutput}");
                    }
                }

                // 處理重新命名 (LibreOffice 固定產出 Test.pdf)
                string expectedFileName = Path.GetFileNameWithoutExtension(excelFilePath) + ".pdf";
                string actualPdfPath = Path.Combine(outputDir, expectedFileName);

                if (File.Exists(actualPdfPath))
                {
                    if (!string.Equals(Path.GetFullPath(actualPdfPath), Path.GetFullPath(pdfOutputPath), StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(pdfOutputPath)) File.Delete(pdfOutputPath);
                        File.Move(actualPdfPath, pdfOutputPath);
                    }
                }

                configManager.LogOperation($"成功將 {Path.GetFileName(excelFilePath)} 轉換為 PDF", LogLevelInfo);
            }
            catch (Exception ex)
            {
                GetConfigManager().LogOperation($"PDF 轉換失敗：{ex.Message}", LogLevelError);
                throw new Exception($"PDF 匯出失敗：{ex.Message}");
            }
        }

        /// <summary>
        /// 匯出為 PDF（指定 LibreOffice 路徑）
        /// </summary>
        public void ExportToPdf(string excelFilePath, string pdfOutputPath, string libreOfficePath)
        {
            try
            {
                // 自動同步記憶體內容磁碟
                if (_workbook != null)
                {
                    string targetSavePath = string.IsNullOrWhiteSpace(excelFilePath) ? _filePath : excelFilePath;
                    if (!string.IsNullOrEmpty(targetSavePath))
                    {
                        SaveWorkbook(targetSavePath);
                        excelFilePath = targetSavePath;
                    }
                }

                if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
                    throw new FileNotFoundException($"找不到 Excel 檔案: {excelFilePath}");

                excelFilePath = Path.GetFullPath(excelFilePath);
                pdfOutputPath = Path.GetFullPath(pdfOutputPath);

                // 驗證 LibreOffice 路徑
                if (!File.Exists(libreOfficePath))
                {
                    throw new FileNotFoundException($"LibreOffice 可執行檔不存在：{libreOfficePath}");
                }

                string outputDir = Path.GetDirectoryName(pdfOutputPath);
                if (string.IsNullOrEmpty(outputDir)) outputDir = Directory.GetCurrentDirectory();

                string rootDir = Path.GetPathRoot(outputDir);
                if (!string.IsNullOrEmpty(rootDir) &&
                    string.Equals(Path.GetFullPath(outputDir), Path.GetFullPath(rootDir), StringComparison.OrdinalIgnoreCase))
                {
                    outputDir = Path.Combine(outputDir, ".");
                }
                else if (outputDir.EndsWith("\\"))
                {
                    outputDir = outputDir.TrimEnd('\\');
                }

                var processInfo = new ProcessStartInfo
                {
                    FileName = libreOfficePath,
                    Arguments = $"--headless --convert-to pdf --outdir \"{outputDir}\" \"{excelFilePath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processInfo))
                {
                    process.WaitForExit(30000);
                    if (process.ExitCode != 0)
                    {
                        string errorOutput = process.StandardError.ReadToEnd();
                        throw new Exception($"轉換失敗：{errorOutput}");
                    }
                }

                // 處理重新命名
                string expectedFileName = Path.GetFileNameWithoutExtension(excelFilePath) + ".pdf";
                string actualPdfPath = Path.Combine(outputDir, expectedFileName);

                if (File.Exists(actualPdfPath) && 
                    !string.Equals(Path.GetFullPath(actualPdfPath), Path.GetFullPath(pdfOutputPath), StringComparison.OrdinalIgnoreCase))
                {
                    if (File.Exists(pdfOutputPath)) File.Delete(pdfOutputPath);
                    File.Move(actualPdfPath, pdfOutputPath);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"PDF 匯出失敗：{ex.Message}");
            }
        }

        /// <summary>
        /// 列印 Excel 檔案（使用 LibreOffice）
        /// </summary>
        public void Print(string excelFilePath, string printerName = null)
        {
            try
            {
                // 自動同步記憶體內容磁碟
                if (_workbook != null)
                {
                    string targetSavePath = string.IsNullOrWhiteSpace(excelFilePath) ? _filePath : excelFilePath;
                    if (!string.IsNullOrEmpty(targetSavePath))
                    {
                        SaveWorkbook(targetSavePath);
                        excelFilePath = targetSavePath;
                    }
                }

                if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
                    throw new FileNotFoundException($"找不到 Excel 檔案: {excelFilePath}");

                excelFilePath = Path.GetFullPath(excelFilePath);

                var configManager = GetConfigManager();
                configManager.LogOperation($"開始列印：{excelFilePath}", LogLevelInfo);

                if (!configManager.IsEnabled())
                {
                    throw new Exception("LibreOffice 列印功能已被停用");
                }

                string libreOfficePath = configManager.GetLibreOfficePath();
                if (string.IsNullOrWhiteSpace(libreOfficePath))
                {
                    throw new Exception("找不到 LibreOffice 路徑。");
                }

                int timeout = configManager.GetTimeoutSeconds();
                string arguments = string.IsNullOrWhiteSpace(printerName)
                    ? $"--headless -p \"{excelFilePath}\""
                    : $"--headless --pt \"{printerName}\" \"{excelFilePath}\"";

                var processInfo = new ProcessStartInfo
                {
                    FileName = libreOfficePath,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(processInfo))
                {
                    bool completed = process.WaitForExit(timeout * 1000);
                    if (!completed)
                    {
                        try { process.Kill(); } catch { }
                        throw new TimeoutException($"LibreOffice 列印超時（{timeout}秒）");
                    }

                    if (process.ExitCode != 0)
                    {
                        string errorOutput = process.StandardError.ReadToEnd();
                        throw new Exception($"LibreOffice 列印失敗 (ExitCode: {process.ExitCode})：{errorOutput}");
                    }
                }

                configManager.LogOperation($"列印命令完成：{Path.GetFileName(excelFilePath)}", LogLevelInfo);
            }
            catch (Exception ex)
            {
                GetConfigManager().LogOperation($"列印失敗：{ex.Message}", LogLevelError);
                throw new Exception($"列印失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// COM 相容的列印方法（供 VBS 呼叫）
        /// </summary>
        public void PrintWorkbook(string excelFilePath, string printerName)
        {
            Print(excelFilePath, string.IsNullOrWhiteSpace(printerName) ? null : printerName);
        }

        /// <summary>
        /// 設定 LibreOffice 執行檔路徑（程式碼優先）
        /// </summary>
        public void SetLibreOfficePath(string libreOfficePath)
        {
            GetConfigManager().SetLibreOfficePathOverride(libreOfficePath);
        }

        /// <summary>
        /// 設定 Log 資料夾路徑（程式碼優先）
        /// </summary>
        public void SetLogDirectory(string logDirectoryPath)
        {
            GetConfigManager().SetLogDirectoryOverride(logDirectoryPath);
        }

        /// <summary>
        /// 設定是否啟用 Log（程式碼優先）
        /// </summary>
        public void SetLogEnabled(bool enabled)
        {
            GetConfigManager().SetLogEnabledOverride(enabled);
        }

        /// <summary>
        /// 設定 Log 等級（程式碼優先）
        /// 1=錯誤, 2=告警, 3=執行狀況
        /// </summary>
        public void SetLogLevel(int logLevel)
        {
            GetConfigManager().SetLogLevelOverride(logLevel);
        }

        #endregion

        #region LibreOffice 配置管理

        /// <summary>
        /// LibreOffice 配置管理類（不依賴設定檔）
        /// </summary>
        private class LibreOfficeConfigManager
        {
            private string _libreOfficePathOverride;
            private string _logDirectoryOverride;
            private bool? _logEnabledOverride;
            private int? _logLevelOverride;

            public void SetLibreOfficePathOverride(string libreOfficePath)
            {
                _libreOfficePathOverride = string.IsNullOrWhiteSpace(libreOfficePath) ? null : libreOfficePath;
            }

            public void SetLogDirectoryOverride(string logDirectoryPath)
            {
                _logDirectoryOverride = string.IsNullOrWhiteSpace(logDirectoryPath) ? null : logDirectoryPath;
            }

            public void SetLogEnabledOverride(bool enabled)
            {
                _logEnabledOverride = enabled;
            }

            public void SetLogLevelOverride(int logLevel)
            {
                if (logLevel < LogLevelError)
                {
                    _logLevelOverride = LogLevelError;
                    return;
                }

                if (logLevel > LogLevelInfo)
                {
                    _logLevelOverride = LogLevelInfo;
                    return;
                }

                _logLevelOverride = logLevel;
            }

            public bool IsEnabled()
            {
                return true;
            }

            public string GetLibreOfficePath()
            {
                if (!string.IsNullOrWhiteSpace(_libreOfficePathOverride) && File.Exists(_libreOfficePathOverride))
                    return _libreOfficePathOverride;

                if (TryFindInSystemPath(out string systemPath) && File.Exists(systemPath))
                    return systemPath;

                string libreOfficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";
                if (File.Exists(libreOfficePath))
                    return libreOfficePath;

                libreOfficePath = @"C:\Program Files (x86)\LibreOffice\program\soffice.exe";
                if (File.Exists(libreOfficePath))
                    return libreOfficePath;

                return null;
            }

            private bool TryFindInSystemPath(out string path)
            {
                path = null;
                try
                {
                    var processInfo = new ProcessStartInfo
                    {
                        FileName = Environment.OSVersion.Platform == PlatformID.Win32NT ? "where" : "which",
                        Arguments = "soffice",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    using (var process = Process.Start(processInfo))
                    {
                        process.WaitForExit();
                        if (process.ExitCode == 0)
                        {
                            path = process.StandardOutput.ReadLine()?.Trim();
                            return !string.IsNullOrWhiteSpace(path);
                        }
                    }
                }
                catch { }

                return false;
            }

            public int GetTimeoutSeconds()
            {
                return 30;
            }

            public void LogOperation(string message, int logLevel)
            {
                try
                {
                    if (!ShouldLog(logLevel))
                        return;

                    string logDirectory = GetLogDirectoryPath();
                    Directory.CreateDirectory(logDirectory);
                    string logPath = Path.Combine(logDirectory, GetLogFileName());
                    string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{GetLogLevelLabel(logLevel)}] {message}";
                    File.AppendAllText(logPath, logMessage + Environment.NewLine);
                }
                catch { }
            }

            public void LogOperation(string message, bool isError = false)
            {
                LogOperation(message, isError ? LogLevelError : LogLevelInfo);
            }

            private bool ShouldLog(int entryLevel)
            {
                if (!IsLoggingEnabled())
                    return false;

                return GetEffectiveLogLevel() >= entryLevel;
            }

            private bool IsLoggingEnabled()
            {
                return _logEnabledOverride ?? true;
            }

            private int GetEffectiveLogLevel()
            {
                return _logLevelOverride ?? LogLevelInfo;
            }

            private string GetLogDirectoryPath()
            {
                if (!string.IsNullOrWhiteSpace(_logDirectoryOverride))
                    return _logDirectoryOverride;

                return Path.GetDirectoryName(typeof(ExcelOperator).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
            }

            private string GetLogFileName()
            {
                AssemblyName assemblyName = typeof(ExcelOperator).Assembly.GetName();
                string dllName = assemblyName.Name;
                string dllVersion = (assemblyName.Version ?? new Version(1, 0, 0, 0)).ToString();
                return $"{dllName}{dllVersion}_{DateTime.Now:yyyyMMdd}.txt";
            }

            private string GetLogLevelLabel(int logLevel)
            {
                if (logLevel <= LogLevelError) return "ERROR";
                if (logLevel == LogLevelWarning) return "WARN";
                return "INFO";
            }
        }

        #endregion

        public void Dispose()
        {
            _workbook?.Dispose();
        }
    }
}

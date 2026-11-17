using System;
using System.Linq;
using System.Runtime.InteropServices;
using Excelopen;
class Program {
 static void Check(bool ok, string name) { if (!ok) throw new Exception(name); Console.WriteLine("PASS " + name); }
 static void Throws(Action a, string name) { try { a(); } catch (Exception) { Console.WriteLine("PASS " + name); return; } throw new Exception(name); }
 static void Main(string[] args) {
  var methods = typeof(IExcelOperator).GetMethods();
  var ids = methods.ToDictionary(x => x.Name, x => x.GetCustomAttributes(typeof(DispIdAttribute), false).Cast<DispIdAttribute>().Single().Value);
  Check(ids.Count == 28 && ids.Values.Distinct().Count() == 28 && ids["SetCellValue"] == 4 && ids["GetAllSheetNames"] == 26 && ids["SetCellFontColor"] == 27 && ids["SetCellFontColorByRC"] == 28, "DispIds unique and unchanged");
  using var op = new ExcelOperator(); op.SetLogEnabled(false); op.CreateNewWorkbook();
  op.SetCellValue("Sheet1", "A1", "red"); op.SetCellValueByRC("Sheet1", 2, 2, "blue");
  op.SetCellFontColor("Sheet1", "A1", 255, 0, 0); op.SetCellFontColorByRC("Sheet1", 2, 2, 0, 0, 255);
  op.SetCellFontColor("Sheet1", "C3", 0, 255, 127);
  Throws(() => op.SetCellFontColor("Sheet1", "A1", -1, 0, 0), "negative RGB rejected");
  Throws(() => op.SetCellFontColorByRC("Sheet1", 2, 2, 0, 256, 0), "RGB above 255 rejected");
  Throws(() => op.SetCellFontColor("missing", "A1", 0, 0, 0), "missing sheet rejected");
  Throws(() => op.SetCellFontColor("Sheet1", "not-a-cell", 0, 0, 0), "invalid address rejected");
  Throws(() => op.SetCellFontColorByRC("Sheet1", 0, 2, 0, 0, 0), "invalid row rejected");
  Check((string)op.GetCellValue("Sheet1", "A1") == "red" && (string)op.GetCellValueByRC("Sheet1", 2, 2) == "blue", "existing get/set");
  op.SaveWorkbook(args[0]); Console.WriteLine("Workbook " + args[0]);
  using var control = new ExcelOperator(); control.SetLogEnabled(false); control.OpenWorkbook(args[1]); control.SaveWorkbook(args[3]);
  using var existing = new ExcelOperator(); existing.SetLogEnabled(false); existing.OpenWorkbook(args[1]); existing.SetCellFontColor("Existing", "A1", 10, 20, 30); existing.SaveWorkbook(args[2]); Console.WriteLine("Existing workbook saved separately");
 }
}

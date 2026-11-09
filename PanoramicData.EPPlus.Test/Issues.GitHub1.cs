using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System.Data;
using System.Globalization;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod, Ignore]
	public void Issue_5()
	{
		var excelFile = new FileInfo(@"c:\temp\bug\test.xlsm");
		using var package = new ExcelPackage(excelFile);
		var ws = package.Workbook.Worksheets.Add("NewWorksheet");
		ws.CodeModule.Code = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n\r\nEnd Sub";
		package.SaveAs(new FileInfo(@"c:\temp\bug\vbafailSaved.xlsm"));
	}
	[TestMethod, Ignore]
	public void Issue_8()
	{

		var l = new List<dynamic>
		{
			1,
			"s"
		};
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("Dynamic Test");
		ws.Cells["a1"].LoadFromCollection(l);
		package.SaveAs(new FileInfo(@"c:\temp\dynamic.xlsx"));


	}
	[TestMethod, Ignore]
	public void Issuer27()
	{
		FileInfo file = new(@"C:\Temp\Test.xlsx");
		var pck = new ExcelPackage(file);
		var ws = pck.Workbook.Worksheets.Add("Worksheet");
		if (ws != null)
		{
			ws.Cells["A1"].Value = "Cell value 1";
			ws.Cells["B1"].Value = "Cell value 2";
			ws.Cells["C1"].Value = "Cell value 3";
			ws.Cells["D1"].Value = "Cell value 4";
			ws.Cells["E1"].Value = "Cell value 5";
		}
		//ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment::Top; // Columns 4 and greater hidden
		ws.Cells.AutoFitColumns(0);
		ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Top; // 4.1.1.0 - exception, 4.0.4.0 - columns 4 and 5 hidden
		ws.Column(4).Hidden = true;
		ws.Column(5).Hidden = true; // span exception
		pck.Save();
	}
	[TestMethod, Ignore]
	public void Issuer26()
	{
		FileInfo file = new(@"C:\Temp\repeatrowcol.xlsx");
		var pck = new ExcelPackage(file);
		var ws = pck.Workbook.Worksheets.Add("Worksheet");
		ws.PrinterSettings.RepeatRows = new ExcelAddress("1:1");
		ws.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");
		pck.Save();
	}
	[TestMethod, Ignore]
	public void Issue32()
	{
		var outputDir = new DirectoryInfo(@"c:\temp\sampleapp");
		var existFile = new FileInfo(Path.Combine(outputDir.FullName, "sample1.xlsx"));
		var newFileName = Path.Combine(outputDir.FullName, "sample1_copied.xlsx");

		File.Copy(existFile.FullName, newFileName, true);
		var newFile = new FileInfo(newFileName);

		using var package = new ExcelPackage(newFile);
		// Add a new worksheet to the empty workbook

		package.Workbook.Worksheets.Add("Inventory_copied", package.Workbook.Worksheets[1]);

		var sourceRange = package.Workbook.Names["row"];
		var sourceRange2 = package.Workbook.Names["roww"];
		var sourceRange3 = package.Workbook.Names["rowww"];
		_ = sourceRange.Worksheet;

		var worksheetTo = package.Workbook.Worksheets["Inventory_copied"];

		var cells = worksheetTo.Cells;
		var rangeTo = cells[1, 1, 1, 16384];
		sourceRange.Copy(rangeTo);
		var rangeTo2 = cells[2, 1, 2, 16384];
		sourceRange2.Copy(rangeTo2);
		var rangeTo3 = cells[4, 1, 4, 16384];
		sourceRange3.Copy(rangeTo3);

		package.Save();
	}
	[TestMethod]
	public void Issue63() // See https://github.com/JanKallman/EPPlus/issues/63
	{
		// Prepare
		var newFile = new FileInfo(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx"));
		try
		{
			using (var package = new ExcelPackage(newFile))
			{
				var ws = package.Workbook.Worksheets.Add("ArrayTest");
				ws.Cells["A1"].Value = 1;
				ws.Cells["A2"].Value = 2;
				ws.Cells["A3"].Value = 3;
				ws.Cells["B1:B3"].CreateArrayFormula("A1:A3");
				package.Save();
			}

			Assert.IsTrue(File.Exists(newFile.FullName));

			// Test: basic support to recognize array formulas after reading Excel workbook file
			using (var package = new ExcelPackage(newFile))
			{
				Assert.AreEqual("A1:A3", package.Workbook.Worksheets["ArrayTest"].Cells["B1"].Formula);
				Assert.IsTrue(package.Workbook.Worksheets["ArrayTest"].Cells["B1"].IsArrayFormula);
			}
		}
		finally
		{
			File.Delete(newFile.FullName);
		}
	}
	[TestMethod, Ignore]
	public void Issue60()
	{
		var ms = new MemoryStream(File.ReadAllBytes(@"c:\temp\sampleapp\sample10\Template.xlsx"));
		using var p = new ExcelPackage();
		p.Load(ms, "");
	}
	[TestMethod]
	public void Issue61()
	{
		DataTable table1 = new("TestTable");
		table1.Columns.Add("name");
		table1.Columns.Add("id");
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("i61");
		ws.Cells["A1"].LoadFromDataTable(table1, true);
	}

	[TestMethod, Ignore]
	public void Issue58()
	{
		var fileInfo = new FileInfo(@"C:\Temp\issue58.xlsx");
		var package = new ExcelPackage(fileInfo);
		if (package.Workbook.Worksheets.Count > 0)
		{
			package.Workbook.Worksheets.Delete("Test");
		}

		var worksheet = package.Workbook.Worksheets.Add("Test");
		worksheet.Cells[1, 1].Value = "Name";
		worksheet.Cells[1, 2].Value = "Address";
		worksheet.Cells[1, 3].Value = "City";
		worksheet.Cells[2, 1].Value = "Esben Rud";
		worksheet.Cells[2, 2].Value = "Enghavevej";
		worksheet.Cells[2, 3].Value = "Odense";
		var r1 = worksheet.ProtectedRanges.Add("Range1", new ExcelAddress(1, 1, 2, 1));
		Console.WriteLine("Range1: " + r1.Name + " " + r1.Address.ToString());
		var r2 = worksheet.ProtectedRanges.Add("Range2", new ExcelAddress(1, 2, 2, 2));
		Console.WriteLine("Range2: " + r2.Name + " " + r2.Address.ToString());
		Console.WriteLine("*Range1: " + r1.Name + " " + r1.Address.ToString());
		Console.WriteLine("*Range2: " + r2.Name + " " + r2.Address.ToString());
		package.Save();
		package.Dispose();
	}
	[TestMethod]
	public void Issue57()
	{
		ExcelPackage pck = new();
		var ws = pck.Workbook.Worksheets.Add("test");
		ws.Cells["A1"].LoadFromArrays([]);
	}
	[TestMethod, Ignore]
	public void Issue55()
	{
		ExcelPackage pck = new(new FileInfo(@"C:\Temp\bug\monocell.xlsx"));
		var ws = pck.Workbook.Worksheets[1];
		Console.WriteLine(ws.Cells["A1"].Text);
	}
	[TestMethod, Ignore]
	public void Issue51()
	{
		var filename = new FileInfo(@"c:\temp\bug\bug51.xlsx");
		using ExcelPackage pck = new(filename);
		var data = pck.Workbook.Worksheets.Add("data");
		data.Cells["A1"].Value = "Product";
		data.Cells["B1"].Value = "Quantity";
		data.Cells["A2"].Value = "Nails";
		data.Cells["B2"].Value = 37;
		data.Cells["A3"].Value = "Hammer";
		data.Cells["B3"].Value = 5;
		data.Cells["A4"].Value = "Saw";
		data.Cells["B4"].Value = 12;

		var dataRange = data.Cells["A1:B4"];

		var pivot = pck.Workbook.Worksheets.Add("pivot");
		var pivotTable = pivot.PivotTables.Add(pivot.Cells["A1"], dataRange, "a&b");
		var tbl = data.Tables.Add(dataRange, "a&c");
		tbl.Name = "_a&c";
		pivotTable.Name = "a&b";
		pck.Save();
	}
	#region Issue 44
	private static readonly string PIVOT_WS_NAME = "Pivot";
	private static readonly string DATA_WS_NAME = "Data";
	[TestMethod, Ignore]
	public void Issue44()
	{
		System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("el-GR");
		using ExcelPackage xlp = new();
		PrepareDoc(xlp);
		GenPivot(xlp);

		var fs = File.Create(@"c:\temp\bug\pivot44.xlsx");
		xlp.SaveAs(fs);
		fs.Close();

	}
	private static void PrepareDoc(ExcelPackage xlp)
	{
		//generate date/value pairs for October 2017
		var series = Enumerable.Range(0, 31);
		var data = from x in series
				   select new { d = new DateTime(2017, 10, x + 1), x };
		//put data in table
		var ws = xlp.Workbook.Worksheets.Add(DATA_WS_NAME);
		var col = 1;
		ws.Cells[1, col++].Value = "Date";
		ws.Cells[1, col].Value = "Value";
		var row = 2;
		foreach (var line in data)
		{
			col = 1;
			ws.Cells[row, col++].Value = line.d;
			ws.Cells[row, col - 1].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
			ws.Cells[row, col].Value = line.x;
			row++;
		}
	}

	private static void GenPivot(ExcelPackage xlp)
	{
		var ws = xlp.Workbook.Worksheets.Add(PIVOT_WS_NAME);
		var srcws = xlp.Workbook.Worksheets[DATA_WS_NAME];
		var piv = ws.PivotTables.Add(ws.Cells[1, 1], srcws.Cells[1, 1, 32, 2], "Pivot1");
		piv.DataFields.Add(piv.Fields["Value"]);
		var dt = piv.RowFields.Add(piv.Fields["Date"]);
		dt.AddDateGrouping(eDateGroupBy.Days | eDateGroupBy.Months);
	}
	#endregion
	[TestMethod]
	public void Issue66()
	{

		using var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Test!");
		ws.Cells["A1"].Value = 1;
		ws.Cells["B1"].Formula = "A1";
		var wb = pck.Workbook;
		wb.Names.Add("Name1", ws.Cells["A1:A2"]);
		ws.Names.Add("Name2", ws.Cells["A1"]);
		pck.Save();
		using var pck2 = new ExcelPackage(pck.Stream);
		_ = pck2.Workbook.Worksheets["Test!"];
	}
}

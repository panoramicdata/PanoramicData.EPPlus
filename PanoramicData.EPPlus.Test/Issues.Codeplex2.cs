using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Table;
using System.Diagnostics;
using System.Reflection;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[Ignore]
	[TestMethod]
	public void Issue15158()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\Output.xlsx"), new FileInfo(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx"));
		var workBook = package.Workbook;
		var worksheet = workBook.Worksheets[1];

		worksheet.DeleteColumn(28);

		if (worksheet.Cells["AA19"].Formula != "")
		{
			throw new InvalidOperationException("this cell should not have formula");
		}

		package.Save();
	}
	[TestMethod]
	public void LoadFromColIssue()
	{
		var l = new List<cls1>
		{
			new() { prop1 = 1 },
			new cls2 { prop1 = 1, prop2 = "test1" }
		};

		var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("Test");

		ws.Cells["A1"].LoadFromCollection(l, true, TableStyles.Light16, BindingFlags.Instance | BindingFlags.Public,
			[typeof(cls2).GetProperty("prop2")]);
	}

	[TestMethod]
	public void Issue15168()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("Test");
		ws.Cells[1, 1].Value = "A1";
		ws.Cells[2, 1].Value = "A2";

		ws.Cells[2, 1].Value = ws.Cells[1, 1].Value;
		Assert.AreEqual("A1", ws.Cells[1, 1].Value);
	}
	[Ignore]
	[TestMethod]
	public void Issue15159()
	{
		var fs = new FileStream(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx", FileMode.OpenOrCreate);
		using (var package = new ExcelPackage(fs))
		{
			package.Save();
		}

		fs.Seek(0, SeekOrigin.Begin);
	}
	[TestMethod]
	public void Issue15179()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("MergeDeleteBug");
		ws.Cells["E3:F3"].Merge = true;
		ws.Cells["E3:F3"].Merge = false;
		ws.DeleteRow(2, 6);
		ws.Cells["A1"].Value = 0;
		_ = ws.Cells["A1"].Value.ToString();
	}
	[Ignore]
	[TestMethod]
	public void Issue15169()
	{
		FileInfo fileInfo = new(@"C:\temp\bug\issue\input.xlsx");

		ExcelPackage excelPackage = new(fileInfo);
		{
			var sheetName = "Labour Costs";

			var ws = excelPackage.Workbook.Worksheets[sheetName];
			excelPackage.Workbook.Worksheets.Delete(ws);

			excelPackage.Workbook.Worksheets.Add(sheetName);

			excelPackage.SaveAs(new FileInfo(@"C:\temp\bug\issue\output2.xlsx"));
		}
	}
	[Ignore]
	[TestMethod]
	public void Issue15172()
	{
		FileInfo fileInfo = new(@"C:\temp\bug\book2.xlsx");

		ExcelPackage excelPackage = new(fileInfo);
		{
			var ws = excelPackage.Workbook.Worksheets[1];

			Assert.AreEqual("IF($R10>=X$2,1,0)", ws.Cells["X10"].Formula);
			ws.Calculate();
			Assert.AreEqual(0D, ws.Cells["X10"].Value);
		}
	}
	[Ignore]
	[TestMethod]
	public void Issue15174()
	{
		using ExcelPackage package = new(new FileInfo(@"C:\temp\bug\MyTemplate.xlsx"));
		package.Workbook.Worksheets[1].Column(2).Style.Numberformat.Format = "dd/mm/yyyy";

		package.SaveAs(new FileInfo(@"C:\temp\bug\MyTemplate2.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void PictureIssue()
	{
		var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("t");
		ws.Drawings.AddPicture("Test", new FileInfo(@"c:\temp\bug\2152228.jpg"));
		p.SaveAs(new FileInfo(@"c:\temp\bug\pic.xlsx"));
	}

	[Ignore]
	[TestMethod]
	public void Issue14988()
	{
		var guid = Guid.NewGuid().ToString("N");
		using var outputStream = new FileStream(@"C:\temp\" + guid + ".xlsx", FileMode.Create);
		using var inputStream = new FileStream(@"C:\temp\bug2.xlsx", FileMode.Open);
		using var package = new ExcelPackage(outputStream, inputStream, "Test");
		var ws = package.Workbook.Worksheets.Add("Test empty");
		ws.Cells["A1"].Value = "Test";
		package.Encryption.Password = "Test2";
		package.Save();
	}

	[TestMethod, Ignore]
	public void Issue15173_1()
	{
		using var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel01.xlsx"));
		var sw = new Stopwatch();
		sw.Start();
		var ws = pck.Workbook.Worksheets.First();
		pck.Workbook.Calculate();
		Assert.AreEqual("20L2300", ws.Cells["F4"].Value);
		Assert.AreEqual("20K2E01", ws.Cells["F5"].Value);
		var f7Val = pck.Workbook.Worksheets["MODELLO-TIPO PANNELLO"].Cells["F7"].Value;
		Assert.AreEqual(13.445419, Math.Round((double)f7Val, 6));
		sw.Stop();
		Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds
	}

	[TestMethod, Ignore]
	public void Issue15173_2()
	{
		using var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel02.xlsx"));
		var sw = new Stopwatch();
		pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
		sw.Start();
		var ws = pck.Workbook.Worksheets.First();
		pck.Workbook.Calculate();
		Assert.AreEqual("20L2300", ws.Cells["F4"].Value);
		Assert.AreEqual("20K2E01", ws.Cells["F5"].Value);
		sw.Stop();
		Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds
	}
	[Ignore]
	[TestMethod]
	public void Issue15154() => Directory.EnumerateFiles(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\").AsParallel().ForAll(file =>
									 {
										 using var package = new ExcelPackage(new FileStream(file, FileMode.Open));
										 package.Workbook.Worksheets[1].Cells[1, 1].Value = file;
										 package.SaveAs(new FileInfo(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\new\" + new FileInfo(file).Name));
									 });
	[Ignore]
	[TestMethod]
	public void Issue15188()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("test");
		worksheet.Column(6).Style.Numberformat.Format = "mm/dd/yyyy";
		worksheet.Column(7).Style.Numberformat.Format = "mm/dd/yyyy";
		worksheet.Column(8).Style.Numberformat.Format = "mm/dd/yyyy";
		worksheet.Column(10).Style.Numberformat.Format = "mm/dd/yyyy";

		worksheet.Cells[2, 6].Value = DateTime.Today;
		var a = worksheet.Cells[2, 6].Text;
		Assert.AreEqual(DateTime.Today.ToString("MM/dd/yyyy"), a);
	}
	[TestMethod, Ignore]
	public void Issue15194()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\i15194-Save.xlsx"), new FileInfo(@"c:\temp\bug\I15194.xlsx"));
		var workBook = package.Workbook;
		var worksheet = workBook.Worksheets[1];

		worksheet.Cells["E3:F3"].Merge = false;

		worksheet.DeleteRow(2, 6);

		package.Save();
	}
	[TestMethod, Ignore]
	public void Issue15195()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"), new FileInfo(@"c:\temp\bug\i15195.xlsx"));
		var workBook = package.Workbook;
		var worksheet = workBook.Worksheets[1];

		worksheet.InsertColumn(8, 2);

		package.Save();
	}
	[TestMethod, Ignore]
	public void Issue14788()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"), new FileInfo(@"c:\temp\bug\GetWorkSheetXmlBad.xlsx"));
		var workBook = package.Workbook;
		var worksheet = workBook.Worksheets[1];

		worksheet.InsertColumn(8, 2);

		package.Save();
	}
	[TestMethod, Ignore]
	public void Issue15167()
	{
		FileInfo fileInfo = new(@"c:\temp\bug\Draw\input.xlsx");

		ExcelPackage excelPackage = new(fileInfo);
		{
			var sheetName = "Board pack";

			var ws = excelPackage.Workbook.Worksheets[sheetName];
			excelPackage.Workbook.Worksheets.Delete(ws);

			ws = excelPackage.Workbook.Worksheets.Add(sheetName);

			excelPackage.SaveAs(new FileInfo(@"c:\temp\bug\output.xlsx"));
		}
	}
}

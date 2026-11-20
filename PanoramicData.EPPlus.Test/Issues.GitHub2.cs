using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod, Ignore]
	public void Issue68()
	{
		using var pck = new ExcelPackage(new FileInfo(@"c:\temp\bug68.xlsx"));
		var ws = pck.Workbook.Worksheets["Sheet1"];
		pck.Workbook.Worksheets.Delete(ws);
		pck.Workbook.Worksheets.Add("Sheet1");
		pck.SaveAs(new FileInfo(@"c:\temp\bug68-2.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue70()
	{
		var documentPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"c:\temp\workbook with comment.xlsx");
		var outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"c:\temp\WorkbookWithCommentOutput.xlsx");
		var fileInfo = new FileInfo(documentPath);
		Assert.IsTrue(fileInfo.Exists);
		using var workbook = new ExcelPackage(fileInfo);
		var ws = workbook.Workbook.Worksheets.First();
		ws.DeleteRow(3); // NRE thrown here
		workbook.SaveAs(new FileInfo(outputPath));
	}

	[TestMethod, Ignore]
	public void Issue100()
	{
		Stream templateFile = new FileStream(@"c:\temp\bug\epplus_drawing_id_issue.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read);
		FileStream outputFile = new(@"c:\temp\bug\epplus_drawing_id_issue_new.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.None);
		using ExcelPackage package = new(templateFile);
		var wb = package.Workbook;
		var sh = wb.Worksheets[1];
		var img_ = File.ReadAllBytes(@"C:\temp\img\background.gif");
		var pic = sh.Drawings.AddPicture("logo", img_);
		pic.SetPosition(1, 1);

		package.SaveAs(outputFile);
	}
	[TestMethod, Ignore]
	public void Issue99()
	{
		var template = @"c:\temp\bug\iss99\Template.xlsx";
		var result = @"c:\temp\bug\iss99\Result.xlsx";
		using var inStream = File.Open(template, FileMode.Open);
		using var outStream = File.Open(result, FileMode.OpenOrCreate, FileAccess.ReadWrite);
		using ExcelPackage xl = new(outStream, inStream);
		xl.Save();
	}
	[TestMethod, Ignore]
	public void Issue94()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\iss94\MergedCellsTemplate.xlsx"));
		var ws = package.Workbook.Worksheets.First();
		package.Workbook.Worksheets.Add("copy", ws);
		package.Workbook.Worksheets.Delete(ws);
		package.SaveAs(new FileInfo(@"c:\temp\bug\iss94\MergedCellsTemplateSaved.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue107()
	{
		using ExcelPackage epIN = new(new FileInfo(@"C:\temp\bug\issue107\in.xlsx"));
		using ExcelPackage epOUT = new(new FileInfo(@"C:\temp\bug\pivotbug107.xlsx"));
		foreach (var sheet in epIN.Workbook.Worksheets)
		{
			epOUT.Workbook.Worksheets.Add(sheet.Name, sheet);
		}

		epIN.Compatibility.IsWorksheets1Based = true;
		epIN.Workbook.Worksheets.Add(epIN.Workbook.Worksheets[1].Name + "-2", epIN.Workbook.Worksheets[1]);
		epIN.Workbook.Worksheets.Add(epIN.Workbook.Worksheets[2].Name + "-2", epIN.Workbook.Worksheets[2]);
		epOUT.Save();
		epIN.SaveAs(new FileInfo(@"C:\temp\bug\pivotbug107-SameWB.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue127()
	{
		using var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\PivotTableTestCase.xlsx"));
		Assert.AreEqual(2, p.Workbook.Worksheets.Count);
		p.SaveAs(new FileInfo(@"C:\temp\bug\PivotTableTestCaseSaved.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue167()
	{
		using var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\test-Errorworkbook.xlsx"));
		Assert.AreEqual(1, p.Workbook.Worksheets.Count);
		p.SaveAs(new FileInfo(@"C:\temp\bug\test-ErrorworkbookSaved.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue155()
	{
		using var pck = new ExcelPackage();
		var ws = pck.Workbook.Worksheets.Add("Sheet1");
		var csv = "A\tB\tC\r\n1\t\t\r\n";
		ws.Cells["A1"].LoadFromText(csv, new ExcelTextFormat { Delimiter = '\t' });
		Assert.IsTrue(ws.Cells["B2"].Value == null);
		var data = pck.GetAsByteArray();
		var path = @"C:\temp\test.xlsx";
		File.WriteAllBytes(path, data);
	}
	[TestMethod, Ignore]
	public void Issue173()
	{
		using ExcelPackage xlPackage = new(new FileInfo($@"C:\temp\bug\issue173.xlsx"));
		var ws = xlPackage.Workbook.Worksheets.First();
		_ = ws.Cells["A4"].Text;
	}
	[TestMethod, Ignore]
	public void Issue176()
	{
		using var pck = new ExcelPackage(new FileInfo($@"C:\temp\bug\issue176.xlsx"));
		Assert.AreEqual(-0.04999M, Math.Round(pck.Workbook.Worksheets[1].Cells["A1"].Style.Fill.BackgroundColor.Tint, 5));
		pck.SaveAs(new FileInfo($@"C:\temp\bug\issue176-saved.xlsx"));
	}

	[TestMethod, Ignore]
	public void Issue178()
	{
		using ExcelPackage package = new();
		var sheet = package.Workbook.Worksheets.Add("TestData");
		var format = new ExcelTextFormat
		{
			TextQualifier = '"'
		};
		var txt = "\"BillingMonth\",\"SequenceNumber\",\"Level7Code\"\r\n";
		txt += "\"022018\",\"1\",\"\"\r\n";

		sheet.Cells["A1"].LoadFromText(txt, format, TableStyles.None, true);
		Assert.AreEqual(null, sheet.Cells["C2"].Value);
	}
	[TestMethod, Ignore]
	public void Issue181()
	{
		using var pck = new ExcelPackage(new FileInfo($@"C:\temp\bug\issue181.xlsx"));
		_ = pck.Workbook.Worksheets.First();
		_ = pck.Workbook.Properties.Author;
		pck.SaveAs(new FileInfo($@"C:\temp\bug\issue181-saved.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue10()
	{
		var fi = new FileInfo($@"C:\temp\bug\issue10.xlsx");
		if (fi.Exists)
		{
			fi.Delete();
		}

		using var pck = new ExcelPackage(fi);
		var ws = pck.Workbook.Worksheets.Add("Pictures");
		var row = 1;
		foreach (var f in Directory.EnumerateFiles(@"c:\temp\addin_temp\Addin\img\open_icon_library-full\icons\ico\16x16\actions\"))
		{
			var b = File.ReadAllBytes(f);
			var pic = ws.Drawings.AddPicture($"Image{(row + 1) / 2}", b);
			pic.SetPosition(row, 0, 0, 0);
			row += 2;
		}

		pck.Save();
	}
	/// <summary>
	/// Creating a new ExcelPackage with an external stream should not dispose of 
	/// that external stream. That is the responsibility of the caller.
	/// Note: This test would pass with EPPlus 4.1.1. In 4.5.1 the line CloseStream() was added
	/// to the ExcelPackage.Dispose() method. That line is redundant with the line before, 
	/// _stream.Close() except that _stream.Close() is only called if the _stream is NOT
	/// an External Stream (and several other conditions).
	/// Note that CloseStream() doesn't do anything different than _stream.Close().
	/// </summary>
	[TestMethod]
	public void Issue184_Disposing_External_Stream()
	{
		// Arrange
		var stream = new MemoryStream();

		using (var excelPackage = new ExcelPackage(stream))
		{
			var worksheet = excelPackage.Workbook.Worksheets.Add("Issue 184");
			worksheet.Cells[1, 1].Value = "Hello EPPlus!";
			excelPackage.SaveAs(stream);
			// Act
		} // This dispose should not dispose of stream.

		// Assert
		Assert.IsTrue(stream.Length > 0);
	}
	[TestMethod]
	public void Issue204()
	{
		using var pack = new ExcelPackage();
		//create sheets
		var sheet1 = pack.Workbook.Worksheets.Add("Sheet 1");
		var sheet2 = pack.Workbook.Worksheets.Add("Sheet 2");
		//set some default values
		sheet1.Cells[1, 1].Value = 1;
		sheet2.Cells[1, 1].Value = 2;
		//fill the formula
		var formula = string.Format("'{0}'!R1C1", sheet1.Name);

		var cell = sheet2.Cells[2, 1];
		cell.FormulaR1C1 = formula;
		//Formula should remain the same
		Assert.AreEqual(formula.ToUpperInvariant(), cell.FormulaR1C1.ToUpperInvariant());
	}
}

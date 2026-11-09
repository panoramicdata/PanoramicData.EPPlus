using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;
using OfficeOpenXml.FormulaParsing;
namespace PanoramicData.EPPlus.Test;

public partial class ReadTemplate
{
	[Ignore]
	[TestMethod]
	public void ReadBug9()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\CovenantsCheckReportTemplate.xlsx"));
		_ = package.Workbook.Worksheets[1];
		package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\new_t.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug10()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\Model_graphes_MBW.xlsm"));

		var ws = package.Workbook.Worksheets["HTTP_data"];
		// Indexed is an int, so an IsNotNull on it could never fail (MSTEST0032); the reads are kept.
		_ = ws.Cells["B4"].Style.Fill.BackgroundColor.Indexed;
		_ = ws.Cells["B5"].Style.Fill.BackgroundColor.Indexed;
	}
	[Ignore]
	[TestMethod]
	public void ReadBug11()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\sample.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		var pck2 = new ExcelPackage();
		pck2.Workbook.Worksheets.Add("Test", ws);
		pck2.SaveAs(new FileInfo(@"c:\temp\SampleNew.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadConditionalFormatting()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\cf2.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		ws.Cells["A1"].Value = 1;
		Assert.AreEqual(eExcelConditionalFormattingRuleType.Equal, ws.ConditionalFormatting[6].Type);
		package.SaveAs(new FileInfo(@"c:\temp\condFormTest.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadStyleBug()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\acquisitions-1993-2.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		ws.Cells["A1"].Value = 1;
		package.SaveAs(new FileInfo(@"c:\temp\condFormTest.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadURL()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\url.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		ws.Cells["A1"].Value = 1;
		package.SaveAs(new FileInfo(@"c:\temp\condFormTest.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadNameError()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\names2.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		ws.Cells["A1"].Value = 1;
		package.SaveAs(new FileInfo(@"c:\temp\TestTableSave.xlsx"));
	}
	[TestMethod, Ignore]
	public void ReadBug12()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\test4.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		ws.Cells["A1"].Value = 1;
		package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug13()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\original.xlsx"));
		_ = package.Workbook.Worksheets[1];
		package.Workbook.Calculate(new OfficeOpenXml.FormulaParsing.ExcelCalculationOption { AllowCirculareReferences = true });
		package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug14()
	{
		var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("Comment");
		ws.Cells["A1"].AddComment("Test av kommentar", "J");
		ws.Comments.RemoveAt(0);
		package.SaveAs(new FileInfo(@"c:\temp\bug\CommentTest.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void ReadBug15()
	{
		var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\ColumnMaxError.xlsx"));
		var ws = package.Workbook.Worksheets[1];
		var col = ws.Column(1);
		col.Style.Fill.PatternType = ExcelFillStyle.Solid;
		col.Style.Fill.BackgroundColor.SetColor(Color.Red);
		package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
	}
}

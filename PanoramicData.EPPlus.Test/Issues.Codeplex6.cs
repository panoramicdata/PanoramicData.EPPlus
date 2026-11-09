using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;

namespace PanoramicData.EPPlus.Test;

public partial class Issues
{
	[TestMethod, Ignore]
	public void Issuer15563()   //And 15562
	{
		using var package = new ExcelPackage();
		var w = package.Workbook.Worksheets.Add("test");
		w.Row(1).Style.Font.Bold = true;
		w.Row(2).Style.Font.Bold = true;

		for (var i = 0; i < 4; i++)
		{
			w.Column(8 + 2 * i).Style.Border.Right.Style = ExcelBorderStyle.Dotted;
		}

		package.SaveAs(new FileInfo(@"c:\temp\bug\stylebug.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issuer15560()
	{
		//Type not set to error when converting shared formula.
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\sharedFormulas.xlsm"));
		package.SaveAs(new FileInfo(@"c:\temp\bug\sharedformulabug.xlsm"));
	}
	[TestMethod, Ignore]
	public void Issuer15558()
	{
		//TODO: ??? works
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\test_file_20161118.xlsx"));
		package.SaveAs(new FileInfo(@"c:\temp\bug\saveproblem.xlsm"));
	}
	/// <summary>
	/// Issue 15561
	/// </summary>
	[TestMethod, Ignore]
	public void Chart_From_Cell_Union_Selector_Bug_Test()
	{
		var existingFile = new FileInfo(@"c:\temp\Chart_From_Cell_Union_Selector_Bug_Test.xlsx");
		if (existingFile.Exists)
		{
			existingFile.Delete();
		}

		using var pck = new ExcelPackage(existingFile);
		var myWorkSheet = pck.Workbook.Worksheets.Add("Content");
		var ExcelWorksheet = pck.Workbook.Worksheets.Add("Chart");

		//Some data
		myWorkSheet.Cells["A1"].Value = "A";
		myWorkSheet.Cells["A2"].Value = 100; myWorkSheet.Cells["A3"].Value = 400; myWorkSheet.Cells["A4"].Value = 200; myWorkSheet.Cells["A5"].Value = 300; myWorkSheet.Cells["A6"].Value = 600; myWorkSheet.Cells["A7"].Value = 500;
		myWorkSheet.Cells["B1"].Value = "B";
		myWorkSheet.Cells["B2"].Value = 300; myWorkSheet.Cells["B3"].Value = 200; myWorkSheet.Cells["B4"].Value = 1000; myWorkSheet.Cells["B5"].Value = 600; myWorkSheet.Cells["B6"].Value = 500; myWorkSheet.Cells["B7"].Value = 200;

		//Pie chart shows with EXTRA B2 entry due to problem with ExcelRange Enumerator
		var values = myWorkSheet.Cells["B2,B4,B6"];  //when the iterator is evaluated it will return the first cell twice: "B2,B2,B4,B6"
		var xvalues = myWorkSheet.Cells["A2,A4,A6"]; //when the iterator is evaluated it will return the first cell twice: "A2,A2,A4,A6"
		var chartBug = ExcelWorksheet.Drawings.AddChart("Chart BAD", eChartType.Pie);
		chartBug.Series.Add(values, xvalues);
		chartBug.Title.Text = "Using ExcelRange";

		//Pie chart shows correctly when using string addresses and avoiding ExcelRange
		var chartGood = ExcelWorksheet.Drawings.AddChart("Chart GOOD", eChartType.Pie);
		chartGood.SetPosition(10, 0, 0, 0);
		chartGood.Series.Add("Content!B2,Content!B4,Content!B6", "Content!A2,Content!A4,Content!A6");
		chartGood.Title.Text = "Using String References";

		pck.Save();
	}
	[TestMethod, Ignore]
	public void Issue15566()
	{
		var TemplateFileName = @"c:\temp\bug\TestWithPivotTablePointingToExcelTableForData.xlsx";
		var ExportFileName = @"c:\temp\bug\TestWithPivotTablePointingToExcelTableForData_Export.xlsx";
		using ExcelPackage excelpackage = new(new FileInfo(TemplateFileName), true);
		excelpackage.SaveAs(new FileInfo(ExportFileName));
	}
	[TestMethod, Ignore]
	public void Issue15564()
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("FormulaBug");
		for (var i = 1; i < 1030; i++)
		{
			ws.Cells[i, 1].Value = i;
			ws.Cells[i, 2].FormulaR1C1 = "rc[-1]+1";
		}

		ws.InsertRow(4, 1025, 3);
		ws.InsertRow(1050, 1025, 3);
		p.SaveAs(new FileInfo(@"c:\temp\bug\fb.xlsx"));
	}
	[TestMethod, Ignore]
	public void Issue15551()    //Works fine?
	{
		using var p = new ExcelPackage();
		var ws = p.Workbook.Worksheets.Add("StyleBug");

		ws.Cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
		ws.Cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
		ws.Cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
		ws.Cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
		p.SaveAs(new FileInfo(@"c:\temp\bug\StyleBug.xlsx"));
	}
	[TestMethod, Ignore]
	public void BugCommentNullAfterRemove()
	{

		var xls = @"c:\temp\bug\in.xlsx";

		ExcelPackage theExcel = new(new FileInfo(xls), true);

		ExcelRangeBase cell;
		ExcelComment cmnt;
		var ws = theExcel.Workbook.Worksheets[1];

		foreach (var addr in "B4 B11".Split(' '))
		{
			cell = ws.Cells[addr];
			cmnt = cell.Comment;

			Assert.IsNotNull(cmnt, "Comment in " + addr + " expected not null ");
			ws.Comments.Remove(cmnt);
		}

	}

	[TestMethod, Ignore]
	public void BugCommentExceptionOnRemove()
	{

		var xls = @"c:\temp\bug\in.xlsx";

		ExcelPackage theExcel = new(new FileInfo(xls), true);

		ExcelRangeBase cell;
		ExcelComment cmnt;
		var ws = theExcel.Workbook.Worksheets[1];

		foreach (var addr in "B4 B16".Split(' '))
		{
			cell = ws.Cells[addr];
			cmnt = cell.Comment;

			try
			{
				ws.Comments.Remove(cmnt);
			}
			catch (Exception ex)
			{
				Assert.Fail("Exception while removing comment at " + addr + ": " + ex.Message);
			}
		}

	}

	[TestMethod]
	public void Issue15548_SumIfsShouldHandleGaps()
	{
		using var package = new ExcelPackage();
		var test = package.Workbook.Worksheets.Add("Test");

		test.Cells["A1"].Value = 1;
		test.Cells["B1"].Value = "A";

		//test.Cells["A2"] is default
		test.Cells["B2"].Value = "A";

		test.Cells["A3"].Value = 1;
		test.Cells["B4"].Value = "B";

		test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

		test.Calculate();

		var result = test.Cells["D2"].GetValue<int>();

		Assert.AreEqual(1, result, string.Format("Expected 1, got {0}", result));
	}
	[TestMethod]
	public void Issue15548_SumIfsShouldHandleBadData()
	{
		using var package = new ExcelPackage();
		var test = package.Workbook.Worksheets.Add("Test");

		test.Cells["A1"].Value = 1;
		test.Cells["B1"].Value = "A";

		test.Cells["A2"].Value = "Not a number";
		test.Cells["B2"].Value = "A";

		test.Cells["A3"].Value = 1;
		test.Cells["B4"].Value = "B";

		test.Cells["D2"].Formula = "SUMIFS(A1:A3,B1:B3,\"A\")";

		test.Calculate();

		var result = test.Cells["D2"].GetValue<int>();

		Assert.AreEqual(1, result, string.Format("Expected 1, got {0}", result));
	}
	[TestMethod, Ignore]
	public void Issue_15585()
	{
		var excelFile = new FileInfo(@"c:\temp\bug\formula_value.xlsx");
		using var package = new ExcelPackage(excelFile);
		// Output from the logger will be written to the following file
		var logfile = new FileInfo(@"C:\temp\EpplusLogFile.txt");
		// Attach the logger before the calculation is performed.
		package.Workbook.FormulaParserManager.AttachLogger(logfile);
		// Calculate - can also be executed on sheet- or range level.
		package.Workbook.Calculate();

		Console.WriteLine(string.Format("Country: \t\t\t{0}", package.Workbook.Worksheets[1].Cells["B1"].Value));
		Console.WriteLine(string.Format("Phone Code - Direct Reference:\t{0}", package.Workbook.Worksheets[1].Cells["B2"].Value.ToString()));
		Console.WriteLine(string.Format("Phone Code - Name Ranges:\t{0}", package.Workbook.Worksheets[1].Cells["B3"].Value.ToString()));
		Console.WriteLine(string.Format("Phone Code - Table reference:\t{0}", package.Workbook.Worksheets[1].Cells["B4"].Value.ToString()));

		// The following method removes any logger attached to the workbook.
		package.Workbook.FormulaParserManager.DetachLogger();
	}
	[TestMethod]
	public void Issue_15641()
	{
		ExcelPackage ep = new();

		var sheet = ep.Workbook.Worksheets.Add("A Sheet");

		sheet.Cells[1, 1].CreateArrayFormula("IF(SUM(B1:J1)>0,SUM(B2:J2))"); //A1
		sheet.Cells[2, 1].Value = sheet.Cells[1, 1].IsArrayFormula; //A2
		sheet.Cells[1, 1].Copy(sheet.Cells[3, 1]); //A3

		Assert.AreEqual(true, sheet.Cells[3, 1].IsArrayFormula);
	}
}

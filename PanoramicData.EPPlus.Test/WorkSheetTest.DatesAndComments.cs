using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using System.Globalization;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	#region Date1904 Test Cases
	private const string Date1904File = "test1904.xlsx";
	private static readonly DateTime Date1904Test1 = new(2008, 2, 29);
	private static readonly DateTime Date1904Test2 = new(1950, 11, 30);

	private static void SaveDate1904TestFile(bool setDate1904)
	{
		if (File.Exists(Date1904File))
		{
			File.Delete(Date1904File);
		}

		ExcelPackage pack = new(new FileInfo(Date1904File));
		if (setDate1904)
		{
			pack.Workbook.Date1904 = true;
		}

		var w = pack.Workbook.Worksheets.Add("test");
		w.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
		w.Cells[1, 1].Value = Date1904Test1;
		w.Cells[2, 1].Value = Date1904Test2;
		pack.Save();
	}

	private static void ResaveDate1904TestFile(bool date1904)
	{
		ExcelPackage pack2 = new(new FileInfo(Date1904File));
		pack2.Workbook.Date1904 = date1904;
		pack2.Save();
	}

	private static void AssertDate1904TestFile(DateTime expected1, DateTime expected2)
	{
		ExcelPackage pack = new(new FileInfo(Date1904File));
		var w = pack.Workbook.Worksheets["test"];

		Assert.AreEqual(expected1, w.Cells[1, 1].Value);
		Assert.AreEqual(expected2, w.Cells[2, 1].Value);
	}

	[TestMethod]
	public void TestDate1904WithoutSetting()
	{
		SaveDate1904TestFile(false);
		AssertDate1904TestFile(Date1904Test1, Date1904Test2);
	}

	[TestMethod]
	public void TestDate1904WithSetting()
	{
		SaveDate1904TestFile(true);
		AssertDate1904TestFile(Date1904Test1, Date1904Test2);
	}

	[TestMethod]
	public void TestDate1904SetAndRemoveSetting()
	{
		SaveDate1904TestFile(true);
		ResaveDate1904TestFile(false);
		AssertDate1904TestFile(Date1904Test1.AddDays(365.5 * -4), Date1904Test2.AddDays(365.5 * -4));
	}

	[TestMethod]
	public void TestDate1904SetAndSetSetting()
	{
		SaveDate1904TestFile(true);
		// Setting the same value again must leave the cells unchanged.
		ResaveDate1904TestFile(true);
		AssertDate1904TestFile(Date1904Test1, Date1904Test2);
	}
	[TestMethod]
	public void ValueText()
	{
		ExcelPackage pck = new();
		var ws = pck.Workbook.Worksheets.Add("TestFormat");
		ws.Cells[1, 1].Value = 25.96;
		ws.Cells[1, 1].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

		_ = ws.Cells[1, 1].Text;
	}
	[TestMethod, Ignore]
	public void SaveToStream()
	{
		var stream = new MemoryStream(File.ReadAllBytes(@"c:\temp\book1.xlsx"));
		var excelPackage = new ExcelPackage(stream);
		excelPackage.Workbook.Worksheets.Add("test");
		excelPackage.Save();
		_ = stream.ToArray();
	}
	[TestMethod, Ignore]
	public void ColumnsTest()
	{
		var excelPackage = new ExcelPackage();
		var ws = excelPackage.Workbook.Worksheets.Add("ColumnTest");


		for (var c = 4; c <= 20; c += 4)
		{
			var col = ws.Column(c);
			col.ColumnMax = c + 3;
		}

		ws.Column(3).Hidden = true;
		ws.Column(6).Hidden = true;
		ws.Column(9).Hidden = true;
		ws.Column(15).Hidden = true;
		ws.Cells["a1:Z1"].Value = "Test";
		ws.Cells["a1:FF33"].AutoFitColumns(0);
		ws.Column(26).ColumnMax = ExcelPackage.MaxColumns;
		excelPackage.SaveAs(new FileInfo(@"c:\temp\autofit.xlsx"));
	}

	[TestMethod]
	public void Comment()
	{
		InitBase();
		var pck = new ExcelPackage();
		var ws1 = pck.Workbook.Worksheets.Add("Comment1");
		ws1.Cells[1, 1].AddComment("Testing", "test1");

		pck.SaveAs(new FileInfo(_worksheetPath + "comment.xlsx"));

		pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
		pck.Compatibility.IsWorksheets1Based = true;
		var ws2 = pck.Workbook.Worksheets[1];
		ws2.Cells[1, 2].AddComment("Testing", "test1");
		pck.Save();
	}

	private void AssertCommentShift(int row2, int col2, Action<ExcelWorksheet> insert, int shiftedRow2, int shiftedCol2)
	{
		InitBase();
		var pck = new ExcelPackage();
		pck.Compatibility.IsWorksheets1Based = true;
		var ws1 = pck.Workbook.Worksheets.Add("Comment1");
		ws1.Cells[3, 3].AddComment("Testing comment 1", "test1");
		ws1.Cells[row2, col2].AddComment("Testing comment 2", "test2");
		var fileInfo = new FileInfo(_worksheetPath + "comment.xlsx");
		pck.SaveAs(fileInfo);
		pck = OpenSavedCommentPackage();
		ws1 = pck.Workbook.Worksheets[1];
		// Ensure the comments were saved in the correct location.
		AssertTwoComments(ws1, row2, col2);
		// Ensure they get shifted.
		insert(ws1);
		AssertTwoComments(ws1, shiftedRow2, shiftedCol2);
		pck.Save();
		pck = OpenSavedCommentPackage();
		ws1 = pck.Workbook.Worksheets[1];
		// Ensure the shifted index is preserved.
		AssertTwoComments(ws1, shiftedRow2, shiftedCol2);
	}

	private ExcelPackage OpenSavedCommentPackage()
	{
		var pck = new ExcelPackage(new FileInfo(_worksheetPath + "comment.xlsx"));
		pck.Compatibility.IsWorksheets1Based = true;
		return pck;
	}

	private static void AssertTwoComments(ExcelWorksheet ws, int row2, int col2)
	{
		Assert.AreEqual("Testing comment 1", ws.Cells[3, 3].Comment.Text);
		Assert.AreEqual("test1", ws.Cells[3, 3].Comment.Author);
		Assert.AreEqual("Testing comment 2", ws.Cells[row2, col2].Comment.Text);
		Assert.AreEqual("test2", ws.Cells[row2, col2].Comment.Author);
	}

	[TestMethod]
	public void CommentShiftsWithRowInserts() => AssertCommentShift(4, 3, ws => ws.InsertRow(4, 4), 8, 3);

	[TestMethod]
	public void CommentShiftsWithColumnInserts() => AssertCommentShift(3, 4, ws => ws.InsertColumn(4, 4), 3, 8);
	#endregion

	[TestMethod]
	public void DateFunctionsWorkWithDifferentCultureDateFormats_US() => AssertDateFunctionsUnderCulture("en-US", "1/15/2014");

	[TestMethod]
	public void DateFunctionsWorkWithDifferentCultureDateFormats_GB() => AssertDateFunctionsUnderCulture("en-GB", "15/1/2014");

	/// <summary>
	/// Runs EDATE on a text date in the given culture on its own thread. Any failure on that thread
	/// is captured and asserted on the test thread: an unhandled exception on a worker thread would
	/// otherwise take the whole test host down. The culture is created without user overrides so
	/// the result does not depend on the machine's regional settings.
	/// </summary>
	private static void AssertDateFunctionsUnderCulture(string cultureName, string textDate)
	{
		var culture = new CultureInfo(cultureName, false);
		var previous = CultureInfo.DefaultThreadCurrentCulture;
		CultureInfo.DefaultThreadCurrentCulture = culture;
		double edate = 0d, eoMonth = 0d;
		Exception failure = null;
		try
		{
			var thread = new Thread(() =>
			{
				try
				{
					Thread.CurrentThread.CurrentCulture = culture;
					using var package = new ExcelPackage();
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = textDate;
					ws.Cells[3, 3].Formula = "EOMONTH(C2, 0)";
					ws.Cells[2, 3].Formula = "EDATE(B2, 0)";
					ws.Calculate();
					edate = Convert.ToDouble(ws.Cells[2, 3].Value);
					eoMonth = Convert.ToDouble(ws.Cells[3, 3].Value);
				}
				catch (Exception ex)
				{
					failure = ex;
				}
			});
			thread.Start();
			thread.Join();
		}
		finally
		{
			CultureInfo.DefaultThreadCurrentCulture = previous;
		}

		Assert.IsNull(failure, failure?.ToString());
		Assert.AreEqual(41654.0, edate);
		Assert.AreEqual(41670.0, eoMonth);
	}
}

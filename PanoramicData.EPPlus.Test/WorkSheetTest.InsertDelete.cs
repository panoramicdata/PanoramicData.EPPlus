using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	//[Ignore]
	//[TestMethod]
	public void InsertDeleteTestRows()
	{
		var ws = _pck.Workbook.Worksheets.Add("InsertDelete");
		ws.Cells["A1:C5"].Value = 1;
		Assert.AreEqual(1, ((object[,])ws.Cells["A1:C5"].Value)[1, 1]);
		ws.Cells["A1:B3"].Merge = true;
		ws.Cells["D3"].Formula = "A2+C5";
		ws.InsertRow(2, 1);

		ws.Cells["A10:C15"].Value = 1;
		ws.Cells["A11:B13"].Merge = true;
		ws.DeleteRow(12, 1, true);

		ws.Cells["a1:B100"].Style.Locked = false;
		ws.Cells["a1:B12"].Style.Hidden = true;
		ws.Protection.IsProtected = true;
		ws.Protection.SetPassword("Password");


		var range = ws.Cells["B2:D100"];

		ApplyPrintAndViewSettings(ws, "B2:D99");

		Assert.AreEqual(2, range.Start.Column);
		Assert.AreEqual(2, range.Start.Row);
		Assert.AreEqual("B2", range.Start.Address);

		Assert.AreEqual(4, range.End.Column);
		Assert.AreEqual(100, range.End.Row);
		Assert.AreEqual("D100", range.End.Address);

		ExcelAddress addr = new("B1:D3");

		Assert.AreEqual(2, addr.Start.Column);
		Assert.AreEqual(1, addr.Start.Row);
		Assert.AreEqual(4, addr.End.Column);
		Assert.AreEqual(3, addr.End.Row);
	}
	//[Ignore]
	//[TestMethod]
	public void InsertDeleteTestColumns()
	{
		var ws = _pck.Workbook.Worksheets.Add("InsertDeleteColumns");
		ws.Cells["A1:C1"].Value = 1;
		ws.Cells["A2:C2"].Value = 2;
		ws.Cells["A3:C3"].Value = 3;
		ws.Cells["A4:C4"].Value = 4;
		ws.Cells["A5:C5"].Value = 5;
		Assert.AreEqual(2, ((object[,])ws.Cells["A1:C5"].Value)[1, 1]);
		ws.Cells["A1:B3"].Merge = true;
		ws.Cells["D3"].Formula = "A2+C5";
		ws.InsertColumn(1, 1);

		ws.Cells["K10:M15"].Value = 1;
		ws.Cells["K11:L13"].Merge = true;
		ws.DeleteColumn(12, 1);

		ws.Cells["X1:Y100"].Style.Locked = false;
		ws.Cells["C1:Y12"].Style.Hidden = true;
		ws.Protection.IsProtected = true;
		ws.Protection.SetPassword("Password");

		ApplyPrintAndViewSettings(ws, "X2:Z99");
	}
	[Ignore]
	[TestMethod]
	public void TestDelete()
	{
		var file = _worksheetPath + "test.xlsx";

		if (File.Exists(file))
			File.Delete(file);

		Create(file);

		ExcelPackage pack = new(new FileInfo(file));
		var w = pack.Workbook.Worksheets["delete"];
		w.DeleteRow(1, 2);

		pack.Save();
	}
	[TestMethod]
	public void InsertRowsSetsOutlineLevel()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		sheet1.Row(15).OutlineLevel = 1;
		sheet1.InsertRow(2, 10, 15);
		for (var i = 2; i < 12; i++)
		{
			Assert.AreEqual(1, sheet1.Row(i).OutlineLevel, $"The outline level of row {i} is not set.");
		}

		Assert.AreEqual(1, sheet1.Row(25).OutlineLevel);
	}

	[TestMethod]
	public void InsertColumnsSetsOutlineLevel()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		sheet1.Column(15).OutlineLevel = 1;
		sheet1.InsertColumn(2, 10, 15);
		for (var i = 2; i < 12; i++)
		{
			Assert.AreEqual(1, sheet1.Column(i).OutlineLevel, $"The outline level of column {i} is not set.");
		}

		Assert.AreEqual(1, sheet1.Column(25).OutlineLevel);
	}

	private static void AssertCopiedRowOutlineLevels(bool crossSheet)
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		sheet1.Row(2).OutlineLevel = 1;
		sheet1.Row(3).OutlineLevel = 1;
		sheet1.Row(4).OutlineLevel = 0;

		var target = crossSheet ? package.Workbook.Worksheets.Add("Sheet2") : sheet1;
		// Set outline levels on rows to be copied over.
		target.Row(6).OutlineLevel = 17;
		target.Row(7).OutlineLevel = 25;
		target.Row(8).OutlineLevel = 29;

		sheet1.Cells["2:4"].Copy(target.Cells["A6"]);
		Assert.AreEqual(1, sheet1.Row(2).OutlineLevel);
		Assert.AreEqual(1, sheet1.Row(3).OutlineLevel);
		Assert.AreEqual(0, sheet1.Row(4).OutlineLevel);

		Assert.AreEqual(1, target.Row(6).OutlineLevel);
		Assert.AreEqual(1, target.Row(7).OutlineLevel);
		Assert.AreEqual(0, target.Row(8).OutlineLevel);
	}

	[TestMethod]
	public void CopyRowSetsOutlineLevelsCorrectly() => AssertCopiedRowOutlineLevels(false);

	[TestMethod]
	public void CopyRowCrossSheetSetsOutlineLevelsCorrectly() => AssertCopiedRowOutlineLevels(true);

	[TestMethod]
	public void CopyColumnSetsOutlineLevelsCorrectly()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		sheet1.Column(2).OutlineLevel = 1;
		sheet1.Column(3).OutlineLevel = 1;
		sheet1.Column(4).OutlineLevel = 0;

		// Set outline levels on rows to be copied over.
		sheet1.Column(6).OutlineLevel = 17;
		sheet1.Column(7).OutlineLevel = 25;
		sheet1.Column(8).OutlineLevel = 29;

		sheet1.Cells["B:D"].Copy(sheet1.Cells["F1"]);
		Assert.AreEqual(1, sheet1.Column(2).OutlineLevel);
		Assert.AreEqual(1, sheet1.Column(3).OutlineLevel);
		Assert.AreEqual(0, sheet1.Column(4).OutlineLevel);

		Assert.AreEqual(1, sheet1.Column(6).OutlineLevel);
		Assert.AreEqual(1, sheet1.Column(7).OutlineLevel);
		Assert.AreEqual(0, sheet1.Column(8).OutlineLevel);
	}
	[TestMethod, Ignore]
	public void Moveissue()
	{
		_pck = new ExcelPackage(new FileInfo(@"C:\temp\bug\FormulaIssue\PreDelete.xlsx"));
		_pck.Workbook.Worksheets[1].DeleteRow(2, 4);
		_pck.SaveAs(new FileInfo(@"c:\temp\move.xlsx"));
	}
	[TestMethod, Ignore]
	public void DelCol()
	{
		_pck = new ExcelPackage(new FileInfo(@"C:\temp\bug\FormulaIssue\PreDeleteCol.xlsx"));
		_pck.Workbook.Worksheets[1].DeleteColumn(5, 1);
		_pck.SaveAs(new FileInfo(@"c:\temp\move.xlsx"));
	}
	[TestMethod, Ignore]
	public void InsCol()
	{
		_pck = new ExcelPackage(new FileInfo(@"C:\temp\bug\FormulaIssue\PreDeleteCol.xlsx"));
		_pck.Workbook.Worksheets[1].InsertColumn(4, 5);
		_pck.SaveAs(new FileInfo(@"c:\temp\move.xlsx"));
	}
	[Ignore]
	[TestMethod]
	public void Deletews()
	{
		FileInfo newFile = new(@"c:\temp\bug\worksheet error.xlsx");
		using (ExcelPackage package = new(newFile))
		{
			package.Workbook.Worksheets.Add("sheet1");
			package.Workbook.Worksheets.Add("sheet2");
			var ws3 = package.Workbook.Worksheets.Add("sheet3");

			package.Workbook.Worksheets.MoveToStart(ws3.Name);
			//save our new workbook and we are done!
			package.Save();
		}

		using (ExcelPackage package = new(newFile))
		{
			package.Workbook.Worksheets.Delete(1);
			package.Workbook.Worksheets.Add("sheet3");
			package.SaveAs(new FileInfo(@"c:\temp\bug\worksheet error_save.xlsx"));
		}
	}
}

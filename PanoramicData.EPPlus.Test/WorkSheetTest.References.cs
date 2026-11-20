using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test;

public partial class WorkSheetTest
{
	//[Ignore]
	//[TestMethod]
	public void CopyRange()
	{
		var ws = _pck.Workbook.Worksheets.Add("CopyTest");

		ws.Cells["A1"].Value = "Single Cell";
		ws.Cells["A2"].Value = "Merged Cells";
		ws.Cells["A2:D30"].Merge = true;
		ws.Cells["A1"].Style.Font.Bold = true;
		ws.Cells["G4:H5"].Merge = true;
		ws.Cells["B3:C5"].Copy(ws.Cells["G4"]);
	}
	//[Ignore]
	//[TestMethod]
	public void CopyMergedRange()
	{
		var ws = _pck.Workbook.Worksheets.Add("CopyMergedRangeTest");

		ws.Cells["A11:C11"].Merge = true;
		ws.Cells["A12:C12"].Merge = true;

		var source = ws.Cells["A11:C12"];
		var target = ws.Cells["A21"];

		source.Copy(target);

		var a21 = ws.Cells[21, 1];
		var a22 = ws.Cells[22, 1];

		Assert.IsTrue(a21.Merge);
		Assert.IsTrue(a22.Merge);
	}

	[TestMethod, Ignore]
	public void Issue15207()
	{
		using ExcelPackage ep = new(new FileInfo(@"c:\temp\bug\worksheet error.xlsx"));
		var wb = ep.Workbook;

		if (wb != null)
		{
			ExcelWorksheet ws = null;

			ws = wb.Worksheets[1];

			//do something with the worksheet
			ws?.Dispose();

			wb.Dispose();

		} //if wb != null

		//do some other things

		//running through this next line now throws the null reference exception
		//so the inbuilt dispose method doesn't work properly.
	}

	[TestMethod]
	public void InsertRowsUpdatesReferencesCorrectly()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		sheet1.Cells[2, 2].Formula = "C3";
		sheet1.Cells[3, 3].Value = "Hello, world!";
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
		sheet1.InsertRow(3, 10);
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet1.Cells[13, 3].Value);
		Assert.AreEqual("C13", sheet1.Cells[2, 2].Formula);
		Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
	}

	[TestMethod]
	public void CrossSheetInsertRowsUpdatesReferencesCorrectly()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
		sheet1.Cells[2, 2].Formula = "Sheet2!C3";
		sheet2.Cells[3, 3].Value = "Hello, world!";
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
		sheet2.InsertRow(3, 10);
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet2.Cells[13, 3].Value);
		Assert.AreEqual("'Sheet2'!C13", sheet1.Cells[2, 2].Formula, true);
		Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
	}

	[TestMethod]
	public void CrossSheetInsertColumnsUpdatesReferencesCorrectly()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
		sheet1.Cells[2, 2].Formula = "'Sheet2'!C3";
		sheet2.Cells[3, 3].Value = "Hello, world!";
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
		sheet2.InsertColumn(3, 10);
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet2.Cells[3, 13].Value);
		Assert.AreEqual("'Sheet2'!M3", sheet1.Cells[2, 2].Formula);
		Assert.AreEqual("Hello, world!", sheet1.Cells[2, 2].Value);
	}

	private static void AssertCrossSheetFormulaAfter(Action<ExcelWorksheet> changeOtherSheet, string expectedFormula)
	{
		FileInfo file = new("report.xlsx");
		using ExcelPackage package = new(file);
		var sheet = package.Workbook.Worksheets.Add("New Sheet");
		var otherSheet = package.Workbook.Worksheets.Add("Other Sheet");
		sheet.Cells[3, 3].Formula = "'Other Sheet'!C3";
		otherSheet.Cells[3, 3].Formula = "45";
		changeOtherSheet(otherSheet);
		Assert.AreEqual(expectedFormula, sheet.Cells[3, 3].Formula);
	}

	[TestMethod]
	public void CrossSheetInsertRowAfterReferencesHasNoEffect()
		=> AssertCrossSheetFormulaAfter(otherSheet => otherSheet.InsertRow(5, 1), "'Other Sheet'!C3");

	[TestMethod]
	public void CrossSheetInsertColumnAfterReferencesHasNoEffect()
		=> AssertCrossSheetFormulaAfter(otherSheet => otherSheet.InsertColumn(5, 1), "'Other Sheet'!C3");

	[TestMethod]
	public void CrossSheetReferenceIsUpdatedWhenSheetIsRenamed()
		=> AssertCrossSheetFormulaAfter(otherSheet => otherSheet.Name = "New Name", "'New Name'!C3");
	private static void AssertCrossSheetCopy(string formula, string expectedCopiedFormula, string expectedCopiedValue)
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
		sheet1.Cells[3, 3].Formula = formula;
		sheet2.Cells[3, 3].Value = "Hello, world!";
		sheet2.Cells[3, 4].Value = "Hello, WORLD!";
		sheet2.Cells[4, 3].Value = "Goodbye, world!";
		sheet2.Cells[4, 4].Value = "Goodbye, WORLD!";
		package.Workbook.Calculate();
		Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
		sheet1.Cells[3, 3].Copy(sheet1.Cells[4, 4]);
		package.Workbook.Calculate();
		Assert.AreEqual(expectedCopiedFormula, sheet1.Cells[4, 4].Formula, true);
		Assert.AreEqual("Hello, world!", sheet1.Cells[3, 3].Value);
		Assert.AreEqual(expectedCopiedValue, sheet1.Cells[4, 4].Value);
	}

	[TestMethod]
	public void CopyCellUpdatesRelativeCrossSheetReferencesCorrectly()
		=> AssertCrossSheetCopy("Sheet2!C3", "'Sheet2'!D4", "Goodbye, WORLD!");

	[TestMethod]
	public void CopyCellUpdatesAbsoluteCrossSheetReferencesCorrectly()
		=> AssertCrossSheetCopy("Sheet2!$C$3", "'Sheet2'!$C$3", "Hello, world!");

	[TestMethod]
	public void CopyCellUpdatesRowAbsoluteCrossSheetReferencesCorrectly()
		=> AssertCrossSheetCopy("Sheet2!C$3", "'Sheet2'!D$3", "Hello, WORLD!");

	[TestMethod]
	public void CopyCellUpdatesColumnAbsoluteCrossSheetReferencesCorrectly()
		=> AssertCrossSheetCopy("Sheet2!$C3", "'Sheet2'!$C4", "Goodbye, world!");

	[TestMethod]
	public void ExcelWorksheetRenameWithStartApostropheThrowsException()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		Assert.ThrowsExactly<ArgumentException>(() => sheet1.Name = "'New Name");
	}

	[TestMethod]
	public void ExcelWorksheetRenameWithEndApostropheThrowsException()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		Assert.ThrowsExactly<ArgumentException>(() => sheet1.Name = "New Name'");
	}
	[TestMethod]
	public void CopySheetWithSharedFormula()
	{
		using ExcelPackage package = new();
		var workbook = package.Workbook;
		var sheet1 = workbook.Worksheets.Add("Sheet1");
		sheet1.Cells[2, 2, 5, 2].Value = new object[,] { { 1 }, { 2 }, { 3 }, { 4 } };
		// Creates a shared formula.
		sheet1.Cells["D2:D5"].Formula = "SUM(B2:C2)";
		var sheet2 = workbook.Worksheets.Copy(sheet1.Name, "Sheet2");
		sheet1.InsertColumn(3, 1);
		// Inserting a column on sheet1 should modify the shared formula on sheet1, but not sheet2.
		Assert.AreEqual("SUM(B2:D2)", sheet1.Cells["E2"].Formula);
		Assert.AreEqual("SUM(B2:C2)", sheet2.Cells["D2"].Formula);
	}
}

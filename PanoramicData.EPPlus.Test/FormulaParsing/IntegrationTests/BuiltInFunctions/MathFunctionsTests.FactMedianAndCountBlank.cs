using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void FactShouldReturnAResult()
	{
		var result = _parser.Parse("Fact(0)");
		Assert.AreEqual(1d, result);
	}

	[TestMethod]
	public void QuotientShouldReturnAResult()
	{
		var result = _parser.Parse("Quotient(5;2)");
		Assert.AreEqual(2, result);
	}

	[TestMethod]
	public void MedianShouldReturnAResult()
	{
		var result = _parser.Parse("Median(1;2;3)");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void CountBlankShouldCalculateEmptyCells()
	{
		Assert.AreEqual(7, CalculateNextToBlankCells("COUNTBLANK(A1:B4)"));
	}

	[TestMethod]
	public void CountBlankShouldCalculateResultOfOffset()
	{
		Assert.AreEqual(1, CalculateNextToBlankCells("COUNTBLANK(OFFSET(A1, 0, 1))"));
	}

	[TestMethod]
	public void DegreesShouldReturnCorrectResult()
	{
		var result = _parser.Parse("DEGREES(0.5)");
		var rounded = Math.Round((double)result, 3);
		Assert.AreEqual(28.648, rounded);
	}

	[TestMethod]
	public void AverateIfsShouldCaluclateResult()
	{
		using var pck = new ExcelPackage();
		var sheet = pck.Workbook.Worksheets.Add("test");
		sheet.Cells["F4"].Value = 1;
		sheet.Cells["F5"].Value = 2;
		sheet.Cells["F6"].Formula = "2 + 2";
		sheet.Cells["F7"].Value = 4;
		sheet.Cells["F8"].Value = 5;

		sheet.Cells["H4"].Value = 3;
		sheet.Cells["H5"].Value = 3;
		sheet.Cells["H6"].Formula = "2 + 2";
		sheet.Cells["H7"].Value = 4;
		sheet.Cells["H8"].Value = 5;

		sheet.Cells["I4"].Value = 2;
		sheet.Cells["I5"].Value = 3;
		sheet.Cells["I6"].Formula = "2 + 2";
		sheet.Cells["I7"].Value = 5;
		sheet.Cells["I8"].Value = 1;

		sheet.Cells["H9"].Formula = "AVERAGEIFS(F4:F8;H4:H8;\">3\";I4:I8;\"<5\")";
		sheet.Calculate();
		Assert.AreEqual(4.5d, sheet.Cells["H9"].Value);
	}

	[TestMethod]
	public void AbsShouldHandleEmptyCell()
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add("test");
		sheet.Cells["A1"].Formula = "ABS(B1)";
		sheet.Calculate();

		Assert.AreEqual(0d, sheet.Cells["A1"].Value);
	}

	private static object CalculateNextToBlankCells(string formula)
	{
		using var pck = new ExcelPackage();
		var sheet = pck.Workbook.Worksheets.Add("test");
		sheet.Cells["A1"].Value = 1;
		sheet.Cells["B2"].Value = string.Empty;
		sheet.Cells["A5"].Formula = formula;
		sheet.Calculate();
		return sheet.Cells["A5"].Value;
	}
}

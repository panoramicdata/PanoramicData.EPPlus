using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges;

[TestClass]
public class MathExcelRangeTests
{
	private ExcelPackage _package;
	private ExcelWorksheet _worksheet;

	[TestInitialize]
	public void Initialize()
	{
		_package = new ExcelPackage();
		_worksheet = _package.Workbook.Worksheets.Add("Test");

		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["A2"].Value = 3;
		_worksheet.Cells["A3"].Value = 6;
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();

	private object Calculate(string address, string formula)
	{
		_worksheet.Cells[address].Formula = formula;
		_worksheet.Calculate();
		return _worksheet.Cells[address].Value;
	}
	[TestMethod]
	public void AbsShouldReturn3() => Assert.AreEqual(3d, Calculate("A4", "ABS(A2)"));

	[TestMethod]
	public void CountShouldReturn3() => Assert.AreEqual(3d, Calculate("A4", "COUNT(A1:A3)"));

	[TestMethod]
	public void CountShouldReturn2IfACellValueIsNull()
	{
		_worksheet.Cells["A2"].Value = null;
		var result = Calculate("A4", "COUNT(A1:A3)");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void CountAShouldReturn3() => Assert.AreEqual(3d, Calculate("A4", "COUNTA(A1:A3)"));

	[TestMethod]
	public void CountIfShouldReturnCorrectResult() => Assert.AreEqual(2d, Calculate("A4", "COUNTIF(A1:A3, \">2\")"));

	[TestMethod]
	public void MaxShouldReturn6() => Assert.AreEqual(6d, Calculate("A4", "Max(A1:A3)"));

	[TestMethod]
	public void MinShouldReturn1() => Assert.AreEqual(1d, Calculate("A4", "Min(A1:A3)"));

	[TestMethod]
	public void AverageShouldReturn3Point333333() => Assert.AreEqual(3d + 1d / 3d, Calculate("A4", "Average(A1:A3)"));

	[TestMethod]
	public void AverageIfShouldHandleSingleRangeNumericExpressionMatch()
	{
		_worksheet.Cells["A4"].Value = "B";
		_worksheet.Cells["A5"].Value = 3;
		Assert.AreEqual(4d, Calculate("A6", "AverageIf(A1:A5,\">1\")"));
	}

	[TestMethod]
	public void AverageIfShouldHandleSingleRangeStringMatch()
	{
		_worksheet.Cells["A4"].Value = "ABC";
		_worksheet.Cells["A5"].Value = "3";
		Assert.AreEqual(4.5d, Calculate("A6", "AverageIf(A1:A5,\">1\")"));
	}

	[TestMethod]
	public void AverageIfShouldHandleLookupRangeStringMatch()
	{
		_worksheet.Cells["A1"].Value = "abc";
		_worksheet.Cells["A2"].Value = "abc";
		_worksheet.Cells["A3"].Value = "def";
		_worksheet.Cells["A4"].Value = "def";
		_worksheet.Cells["A5"].Value = "abd";

		_worksheet.Cells["B1"].Value = 1;
		_worksheet.Cells["B2"].Value = 3;
		_worksheet.Cells["B3"].Value = 5;
		_worksheet.Cells["B4"].Value = 6;
		_worksheet.Cells["B5"].Value = 7;

		Assert.AreEqual(2d, Calculate("A6", "AverageIf(A1:A5,\"abc\",B1:B5)"));
	}

	[TestMethod]
	public void AverageIfShouldHandleLookupRangeStringNumericMatch()
	{
		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["A2"].Value = 3;
		_worksheet.Cells["A3"].Value = 3;
		_worksheet.Cells["A4"].Value = 5;
		_worksheet.Cells["A5"].Value = 2;

		_worksheet.Cells["B1"].Value = 3;
		_worksheet.Cells["B2"].Value = 3;
		_worksheet.Cells["B3"].Value = 2;
		_worksheet.Cells["B4"].Value = 1;
		_worksheet.Cells["B5"].Value = 8;

		Assert.AreEqual(2d, Calculate("A6", "AverageIf(A1:A5,\">2\",B1:B5)"));
	}

	[TestMethod]
	public void AverageIfShouldHandleLookupRangeStringWildCardMatch()
	{
		_worksheet.Cells["A1"].Value = "abc";
		_worksheet.Cells["A2"].Value = "abc";
		_worksheet.Cells["A3"].Value = "def";
		_worksheet.Cells["A4"].Value = "def";
		_worksheet.Cells["A5"].Value = "abd";

		_worksheet.Cells["B1"].Value = 1;
		_worksheet.Cells["B2"].Value = 3;
		_worksheet.Cells["B3"].Value = 5;
		_worksheet.Cells["B4"].Value = 6;
		_worksheet.Cells["B5"].Value = 8;

		Assert.AreEqual(4d, Calculate("A6", "AverageIf(A1:A5, \"ab*\",B1:B5)"));
	}

	[TestMethod]
	public void SumProductWithRange()
	{
		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["A2"].Value = 2;
		_worksheet.Cells["A3"].Value = 3;
		_worksheet.Cells["B1"].Value = 5;
		_worksheet.Cells["B2"].Value = 6;
		_worksheet.Cells["B3"].Value = 4;
		var result = Calculate("A4", "SUMPRODUCT(A1:A3,B1:B3)");
		Assert.AreEqual(29d, result);
	}

	[TestMethod]
	public void SumProductWithRangeAndValues()
	{
		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["A2"].Value = 2;
		_worksheet.Cells["A3"].Value = 3;
		_worksheet.Cells["B1"].Value = 5;
		_worksheet.Cells["B2"].Value = 6;
		_worksheet.Cells["B3"].Value = 4;
		var result = Calculate("A4", "SUMPRODUCT(A1:A3,B1:B3,{2,4,1})");
		Assert.AreEqual(70d, result);
	}

	[TestMethod]
	public void SignShouldReturn1WhenRefIsPositive() => Assert.AreEqual(1d, Calculate("A4", "SIGN(A1)"));

	[TestMethod]
	public void SubTotalShouldNotIncludeHiddenRow()
	{
		_worksheet.Row(2).Hidden = true;
		var result = Calculate("A4", "SUBTOTAL(109,A1:A3)");
		Assert.AreEqual(7d, result);
	}

	[TestMethod]
	public void SumProductShouldWorkWithSingleCellArray()
	{
		_worksheet.Cells["A1"].Value = 1;
		_worksheet.Cells["A2"].Value = 2;
		var result = Calculate("A4", "SUMPRODUCT(A1:A1, A2:A2)");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void ShouldIgnoreNullValues() => Assert.AreEqual(0d, Calculate("B3", "C4 + D4"));
}

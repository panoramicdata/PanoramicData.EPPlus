
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges;

public partial class MathExcelRangeTests
{
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
		OneThreeSixSheet.SetColumn(_worksheet, "A", "abc", "abc", "def", "def", "abd");

		OneThreeSixSheet.SetColumn(_worksheet, "B", 1, 3, 5, 6, 7);

		Assert.AreEqual(2d, Calculate("A6", "AverageIf(A1:A5,\"abc\",B1:B5)"));
	}

	[TestMethod]
	public void AverageIfShouldHandleLookupRangeStringNumericMatch()
	{
		OneThreeSixSheet.SetColumn(_worksheet, "A", 1, 3, 3, 5, 2);

		OneThreeSixSheet.SetColumn(_worksheet, "B", 3, 3, 2, 1, 8);

		Assert.AreEqual(2d, Calculate("A6", "AverageIf(A1:A5,\">2\",B1:B5)"));
	}

	[TestMethod]
	public void AverageIfShouldHandleLookupRangeStringWildCardMatch()
	{
		OneThreeSixSheet.SetColumn(_worksheet, "A", "abc", "abc", "def", "def", "abd");

		OneThreeSixSheet.SetColumn(_worksheet, "B", 1, 3, 5, 6, 8);

		Assert.AreEqual(4d, Calculate("A6", "AverageIf(A1:A5, \"ab*\",B1:B5)"));
	}

	[TestMethod]
	public void SumProductWithRange()
	{
		OneThreeSixSheet.SetColumn(_worksheet, "A", 1, 2, 3);
		OneThreeSixSheet.SetColumn(_worksheet, "B", 5, 6, 4);
		var result = Calculate("A4", "SUMPRODUCT(A1:A3,B1:B3)");
		Assert.AreEqual(29d, result);
	}

	[TestMethod]
	public void SumProductWithRangeAndValues()
	{
		OneThreeSixSheet.SetColumn(_worksheet, "A", 1, 2, 3);
		OneThreeSixSheet.SetColumn(_worksheet, "B", 5, 6, 4);
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

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

public partial class RefAndLookupTests
{
	private static CompileResult MatchInFirstRow(object value, int matchType, int a1, int b1, int c1)
		=> ExecuteOnFakeCells(new Match(), [value, "A1:C1", matchType], false, (1, 1, a1), (1, 2, b1), (1, 3, c1));

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
	{
		var result = MatchInFirstRow(3, 0, 1, 3, 5);
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValVertical_MatchTypeExact()
	{
		var result = ExecuteOnFakeCells(new Match(), [3, "A1:A3", 0], true,
			(1, 1, 1), (2, 1, 3), (3, 1, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestBelow()
	{
		var result = MatchInFirstRow(4, 1, 1, 3, 5);
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
	{
		var result = MatchInFirstRow(6, -1, 10, 8, 5);
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
	{
		var result = MatchInFirstRow(10, -1, 10, 8, 5);
		Assert.AreEqual(1, result.Result);
	}

	[TestMethod]
	public void MatchShouldHandleAddressOnOtherSheet()
	{
		using var package = new ExcelPackage();
		var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
		var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
		sheet1.Cells["A1"].Formula = "Match(10, Sheet2!A1:Sheet2!A3, 0)";
		sheet2.Cells["A1"].Value = 9;
		sheet2.Cells["A2"].Value = 10;
		sheet2.Cells["A3"].Value = 11;
		sheet1.Calculate();
		Assert.AreEqual(2, sheet1.Cells["A1"].Value);
	}
}

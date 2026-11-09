using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

public partial class RefAndLookupTests
{
	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
	{
		var result = ExecuteOnFakeCells(new Match(), [3, "A1:C1", 0], false,
			(1, 1, 1), (1, 2, 3), (1, 3, 5));
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
		var result = ExecuteOnFakeCells(new Match(), [4, "A1:C1", 1], false,
			(1, 1, 1), (1, 2, 3), (1, 3, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
	{
		var result = ExecuteOnFakeCells(new Match(), [6, "A1:C1", -1], false,
			(1, 1, 10), (1, 2, 8), (1, 3, 5));
		Assert.AreEqual(2, result.Result);
	}

	[TestMethod]
	public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
	{
		var result = ExecuteOnFakeCells(new Match(), [10, "A1:C1", -1], false,
			(1, 1, 10), (1, 2, 8), (1, 3, 5));
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

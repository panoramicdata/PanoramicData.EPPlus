using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class SumIfTests : CriteriaRangeTestBase
{
	[TestMethod]
	public void SumIfNumeric() => Assert.AreEqual(8d, Criteria.ExecuteWithSumRange(new SumIf(), 1d, 2d, 3d, ">1"));

	[TestMethod]
	public void SumIfNonNumeric() => Assert.AreEqual(8d, Criteria.ExecuteWithSumRange(new SumIf(), "Monday", "Tuesday", "Thursday", "T*day"));

	[TestMethod]
	public void SumIfNumericExpression() => Assert.AreEqual(1d, Criteria.Execute(new SumIf(), null, 1d, "Not Empty", 1d));

	[TestMethod]
	public void SumIfEqualToEmptyString() => Assert.AreEqual(1d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", ""));

	[TestMethod]
	public void SumIfNotEqualToNull() => Assert.AreEqual(8d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", "<>"));

	[TestMethod]
	public void SumIfEqualToZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, 0d, "0"));

	[TestMethod]
	public void SumIfNotEqualToZero() => Assert.AreEqual(4d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, 0d, "<>0"));

	[TestMethod]
	public void SumIfGreaterThanZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, 1d, ">0"));

	[TestMethod]
	public void SumIfGreaterThanOrEqualToZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, 1d, ">=0"));

	[TestMethod]
	public void SumIfLessThanZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, -1d, "<0"));

	[TestMethod]
	public void SumIfLessThanOrEqualToZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, -1d, "<=0"));

	[TestMethod]
	public void SumIfLessThanCharacter() => Assert.AreEqual(3d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", "<a"));

	[TestMethod]
	public void SumIfLessThanOrEqualToCharacter() => Assert.AreEqual(3d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", "<=a"));

	[TestMethod]
	public void SumIfGreaterThanCharacter() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", ">a"));

	[TestMethod]
	public void SumIfGreaterThanOrEqualToCharacter() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", ">=a"));

	[TestMethod]
	public void SumIfHandleDates() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new SumIf(), null, string.Empty, "Not Empty", ">=a"));

	[TestMethod]
	public void SumIfShouldHandleBooleanArg()
	{
		using var pck = new ExcelPackage();
		var sheet = pck.Workbook.Worksheets.Add("test");
		sheet.Cells["A1"].Value = true;
		sheet.Cells["B1"].Value = 1;
		sheet.Cells["A2"].Value = false;
		sheet.Cells["B2"].Value = 1;
		sheet.Cells["C1"].Formula = "SUMIF(A1:A2,TRUE,B1:B2)";
		sheet.Calculate();
		Assert.AreEqual(1d, sheet.Cells["C1"].Value);
	}
}

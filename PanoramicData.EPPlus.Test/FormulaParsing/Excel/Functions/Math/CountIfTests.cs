using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class CountIfTests : CriteriaRangeTestBase
{
	[TestMethod]
	public void CountIfNumeric() => Assert.AreEqual(2d, Criteria.Execute(new CountIf(), 1d, 2d, 3d, ">1"));

	[TestMethod]
	public void CountIfNonNumeric() => Assert.AreEqual(2d, Criteria.Execute(new CountIf(), "Monday", "Tuesday", "Thursday", "T*day"));

	public void CountIfNullExpression()
	{
		Worksheet.Cells["A1"].Value = null;
		Worksheet.Cells["A2"].Value = 1d;
		Worksheet.Cells["A3"].Value = null;
		Worksheet.Cells["B2"].Value = null;
		var func = new CountIf();
		var range1 = Provider.GetRange(Worksheet.Name, 1, 1, 3, 1);
		var range2 = Provider.GetRange(Worksheet.Name, 2, 2, 2, 2);
		var args = FunctionsHelper.CreateArgs(range1, range2);
		var result = func.Execute(args, Context);
		Assert.AreEqual(0d, result.Result);
	}

	[TestMethod]
	public void CountIfNumericExpression() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, 1d, "Not Empty", 1d));

	[TestMethod]
	public void CountIfEqualToEmptyString() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", ""));

	[TestMethod]
	public void CountIfNotEqualToNull() => Assert.AreEqual(2d, Criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", "<>"));

	[TestMethod]
	public void CountIfEqualToZero() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, 0d, "0"));

	[TestMethod]
	public void CountIfNotEqualToZero() => Assert.AreEqual(2d, Criteria.Execute(new CountIf(), null, string.Empty, 0d, "<>0"));

	[TestMethod]
	public void CountIfGreaterThanZero() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, 1d, ">0"));

	[TestMethod]
	public void CountIfGreaterThanOrEqualToZero() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, 1d, ">=0"));

	[TestMethod]
	public void CountIfLessThanZero() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, -1d, "<0"));

	[TestMethod]
	public void CountIfLessThanOrEqualToZero() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, -1d, "<=0"));

	[TestMethod]
	public void CountIfLessThanCharacter() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", "<a"));

	[TestMethod]
	public void CountIfLessThanOrEqualToCharacter() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", "<=a"));

	[TestMethod]
	public void CountIfGreaterThanCharacter() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", ">a"));

	[TestMethod]
	public void CountIfGreaterThanOrEqualToCharacter() => Assert.AreEqual(1d, Criteria.Execute(new CountIf(), null, string.Empty, "Not Empty", ">=a"));
}

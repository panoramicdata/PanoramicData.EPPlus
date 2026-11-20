using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class AverageIfTests : CriteriaRangeTestBase
{
	[TestMethod]
	public void AverageIfNumeric() => Assert.AreEqual(4d, Criteria.ExecuteWithSumRange(new AverageIf(), 1d, 2d, 3d, ">1"));

	[TestMethod]
	public void AverageIfNonNumeric() => Assert.AreEqual(4d, Criteria.ExecuteWithSumRange(new AverageIf(), "Monday", "Tuesday", "Thursday", "T*day"));

	[TestMethod]
	public void AverageIfNumericExpression() => Assert.AreEqual(1d, Criteria.Execute(new AverageIf(), null, 1d, "Not Empty", 1d));

	[TestMethod]
	public void AverageIfEqualToEmptyString() => Assert.AreEqual(1d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, "Not Empty", ""));

	[TestMethod]
	public void AverageIfNotEqualToNull() => Assert.AreEqual(4d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, "Not Empty", "<>"));

	[TestMethod]
	public void AverageIfEqualToZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, 0d, "0"));

	[TestMethod]
	public void AverageIfNotEqualToZero() => Assert.AreEqual(2d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, 0d, "<>0"));

	[TestMethod]
	public void AverageIfGreaterThanZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, 1d, ">0"));

	[TestMethod]
	public void AverageIfGreaterThanOrEqualToZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, 1d, ">=0"));

	[TestMethod]
	public void AverageIfLessThanZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, -1d, "<0"));

	[TestMethod]
	public void AverageIfLessThanOrEqualToZero() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, -1d, "<=0"));

	[TestMethod]
	public void AverageIfLessThanCharacter() => Assert.AreEqual(3d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, "Not Empty", "<a"));

	[TestMethod]
	public void AverageIfLessThanOrEqualToCharacter() => Assert.AreEqual(3d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, "Not Empty", "<=a"));

	[TestMethod]
	public void AverageIfGreaterThanCharacter() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, "Not Empty", ">a"));

	[TestMethod]
	public void AverageIfGreaterThanOrEqualToCharacter() => Assert.AreEqual(5d, Criteria.ExecuteWithSumRange(new AverageIf(), null, string.Empty, "Not Empty", ">=a"));
}

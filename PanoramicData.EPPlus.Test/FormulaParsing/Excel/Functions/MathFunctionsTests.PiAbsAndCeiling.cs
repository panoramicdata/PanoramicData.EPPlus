using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void PiShouldReturnPIConstant() => Assert.AreEqual(System.Math.Round(System.Math.PI, 14), ExecuteFunction(new Pi(), 0));

	[TestMethod]
	public void AbsShouldReturnCorrectResult() => Assert.AreEqual(3d, ExecuteFunction(new Abs(), -3d));

	[TestMethod]
	public void AsinShouldReturnCorrectResult() => Assert.AreEqual(1.5708, ExecuteFunctionRounded(new Asin(), 4, 1d));

	[TestMethod]
	public void AsinhShouldReturnCorrectResult() => Assert.AreEqual(0.0998, ExecuteFunctionRounded(new Asinh(), 4, 0.1d));

	[TestMethod]
	public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0() => Assert.AreEqual(22.36d, ExecuteFunction(new Ceiling(), 22.35d, 0.01));

	[TestMethod]
	public void CeilingBugTest1() => Assert.AreEqual(100d, ExecuteFunction(new Ceiling(), 100d, 100d));
	[TestMethod]
	public void CeilingBugTest2() => Assert.AreEqual(12000d, ExecuteFunction(new Ceiling(), 12000d, 1000d));

	[TestMethod]
	public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
	{
		var expectedValue = -22.4d;
		var func = new Ceiling();
		var args = FunctionsHelper.CreateArgs(-22.35d, -0.1);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
	}

	[TestMethod]
	public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1() => Assert.AreEqual(23d, ExecuteFunction(new Ceiling(), 22.35d, 1));

	[TestMethod]
	public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10() => Assert.AreEqual(30d, ExecuteFunction(new Ceiling(), 22.35d, 10));

	[TestMethod]
	public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative() => Assert.AreEqual(-30d, ExecuteFunction(new Ceiling(), -22.35d, -10));

	[TestMethod]
	public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
	{
		var func = new Ceiling();
		var args = FunctionsHelper.CreateArgs(22.35d, -1);
		Assert.ThrowsExactly<InvalidOperationException>(() => func.Execute(args, _parsingContext));
	}
}

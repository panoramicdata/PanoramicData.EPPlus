using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void RoundShouldReturnCorrectResult() => Assert.AreEqual(2.343d, ExecuteFunction(new Round(), 2.3433, 3));

	[TestMethod]
	public void RoundShouldReturnCorrectResultWhenNbrOfDecimalsIsNegative() => Assert.AreEqual(9000d, ExecuteFunction(new Round(), 9333, -3));

	[TestMethod]
	public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1() => Assert.AreEqual(26.7d, ExecuteFunction(new Floor(), 26.75d, 0.1));

	[TestMethod]
	public void FloorShouldReturnCorrectResultWhenSignificanceIs1() => Assert.AreEqual(26d, ExecuteFunction(new Floor(), 26.75d, 1));

	[TestMethod]
	public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1() => Assert.AreEqual(-26d, ExecuteFunction(new Floor(), -26.75d, -1));
	[TestMethod]
	public void FloorBugTest1() => Assert.AreEqual(100d, ExecuteFunction(new Floor(), 100d, 100d));
	[TestMethod]
	public void FloorBugTest2() => Assert.AreEqual(12000d, ExecuteFunction(new Floor(), 12000d, 1000d));

	[TestMethod]
	public void RandShouldReturnAValueBetween0and1()
	{
		var func = new Rand();
		var args = Array.Empty<FunctionArgument>();
		var result1 = func.Execute(args, _parsingContext);
		Assert.IsTrue((double)result1.Result is > 0 and < 1);
		var result2 = func.Execute(args, _parsingContext);
		Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
		Assert.IsTrue((double)result2.Result is > 0 and < 1);
	}

	[TestMethod]
	public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
	{
		var func = new RandBetween();
		var args = FunctionsHelper.CreateArgs(1, 5);
		var result = func.Execute(args, _parsingContext);
		CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
	}

	[TestMethod]
	public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
	{
		var func = new RandBetween();
		var args = FunctionsHelper.CreateArgs(-5, 0);
		var result = func.Execute(args, _parsingContext);
		CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
	}
}

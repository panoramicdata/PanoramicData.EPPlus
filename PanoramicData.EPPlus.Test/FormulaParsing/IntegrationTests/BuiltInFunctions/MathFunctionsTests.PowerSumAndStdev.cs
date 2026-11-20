
namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void PowerShouldReturnCorrectResult()
	{
		var result = _parser.Parse("Power(3, 3)");
		Assert.AreEqual(27d, result);
	}

	[TestMethod]
	public void SqrtShouldReturnCorrectResult()
	{
		var result = _parser.Parse("sqrt(9)");
		Assert.AreEqual(3d, result);
	}

	[TestMethod]
	public void PiShouldReturnCorrectResult()
	{
		var expectedValue = Math.Round(Math.PI, 14);
		var result = _parser.Parse("Pi()");
		Assert.AreEqual(expectedValue, result);
	}

	[TestMethod]
	public void CeilingShouldReturnCorrectResult()
	{
		var expectedValue = 22.4d;
		var result = _parser.Parse("ceiling(22.35, 0.1)");
		Assert.AreEqual(expectedValue, result);
	}

	[TestMethod]
	public void FloorShouldReturnCorrectResult()
	{
		var expectedValue = 22.3d;
		var result = _parser.Parse("Floor(22.35, 0.1)");
		Assert.AreEqual(expectedValue, result);
	}

	[TestMethod]
	public void SumShouldReturnCorrectResultWithInts()
	{
		var result = _parser.Parse("sum(1, 2)");
		Assert.AreEqual(3d, result);
	}

	[TestMethod]
	public void SumShouldReturnCorrectResultWithDecimals()
	{
		var result = _parser.Parse("sum(1,2.5)");
		Assert.AreEqual(3.5d, result);
	}

	[TestMethod]
	public void SumShouldReturnCorrectResultWithEnumerable()
	{
		var result = _parser.Parse("sum({1;2;3;-1}, 2.5)");
		Assert.AreEqual(7.5d, result);
	}

	[TestMethod]
	public void SumsqShouldReturnCorrectResultWithEnumerable()
	{
		var result = _parser.Parse("sumsq({2;3})");
		Assert.AreEqual(13d, result);
	}

	[TestMethod]
	public void SubtotalShouldNegateExpression()
	{
		var result = _parser.Parse("-subtotal(2;{1;2})");
		Assert.AreEqual(-2d, result);
	}

	[TestMethod]
	public void StdevShouldReturnAResult()
	{
		var result = _parser.Parse("stdev(1;2;3;4)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void StdevPShouldReturnAResult()
	{
		var result = _parser.Parse("stdevp(2,3,4)");
		Assert.IsInstanceOfType<double>(result);
	}
}

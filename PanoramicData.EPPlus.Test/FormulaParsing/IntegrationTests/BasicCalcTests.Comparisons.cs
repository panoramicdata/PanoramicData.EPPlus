

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests;

public partial class BasicCalcTests
{
	[TestMethod]
	public void ThreeGreaterThanTwoShouldBeTrue()
	{
		var result = _parser.Parse("3 > 2");
		Assert.IsTrue((bool)result);
	}

	[TestMethod]
	public void ThreeLessThanTwoShouldBeFalse()
	{
		var result = _parser.Parse("3 < 2");
		Assert.IsFalse((bool)result);
	}

	[TestMethod]
	public void ThreeLessThanOrEqualToThreeShouldBeTrue()
	{
		var result = _parser.Parse("3 <= 3");
		Assert.IsTrue((bool)result);
	}

	[TestMethod]
	public void ThreeLessThanOrEqualToTwoDotThreeShouldBeFalse()
	{
		var result = _parser.Parse("3 <= 2.3");
		Assert.IsFalse((bool)result);
	}

	[TestMethod]
	public void ThreeGreaterThanOrEqualToThreeShouldBeTrue()
	{
		var result = _parser.Parse("3 >= 3");
		Assert.IsTrue((bool)result);
	}

	[TestMethod]
	public void TwoDotTwoGreaterThanOrEqualToThreeShouldBeFalse()
	{
		var result = _parser.Parse("2.2 >= 3");
		Assert.IsFalse((bool)result);
	}

	[TestMethod]
	public void TwelveAndTwelveShouldBeEqual()
	{
		var result = _parser.Parse("2=2");
		Assert.IsTrue((bool)result);
	}

	[TestMethod]
	public void TenPercentShouldBe0Point1()
	{
		var result = _parser.Parse("10%");
		Assert.AreEqual(0.1, result);
	}

	[TestMethod]
	public void ShouldHandleMultiplePercentSigns()
	{
		var result = _parser.Parse("10%%");
		Assert.AreEqual(0.001, result);
	}

	[TestMethod]
	public void ShouldHandlePercentageOnFunctionResult()
	{
		var result = _parser.Parse("SUM(1;2;3)%");
		Assert.AreEqual(0.06, result);
	}

	[TestMethod]
	public void ShouldHandlePercentageOnParantethis()
	{
		var result = _parser.Parse("(1+2)%");
		Assert.AreEqual(0.03, result);
	}

	[TestMethod]
	public void ShouldIgnoreLeadingPlus()
	{
		var result = _parser.Parse("+(1-2)");
		Assert.AreEqual(-1d, result);
	}

	[TestMethod]
	public void ShouldHandleDecimalNumberWhenDividingIntegers()
	{
		var result = _parser.Parse("224567455/400000000*500000");
		Assert.AreEqual(280709.31875, result);
	}

	[TestMethod]
	public void ShouldNegateExpressionInParenthesis()
	{
		var result = _parser.Parse("-(1+2)");
		Assert.AreEqual(-3d, result);
	}
}

using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void CountShouldReturnNumberOfNumericItems() => Assert.AreEqual(5d, ExecuteFunction(new Count(), 1d, 2m, 3, new DateTime(2012, 4, 1), "4"));

	[TestMethod]
	public void CountShouldIncludeNumericStringsAndDatesInArray() => Assert.AreEqual(5d, ExecuteFunction(new Count(), FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4")));


	[TestMethod]
	public void CountShouldIncludeEnumerableMembers() => Assert.AreEqual(3d, ExecuteFunction(new Count(), 1d, FunctionsHelper.CreateArgs(12, 13)));

	[TestMethod, Ignore]
	public void CountShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
	{
		var result = ExecuteWithHiddenArg(new Count(), a => a.ElementAt(0), 1d, FunctionsHelper.CreateArgs(12, 13));
		Assert.AreEqual(2d, result.Result);
	}

	[TestMethod]
	public void CountAShouldReturnNumberOfNonWhitespaceItems() => Assert.AreEqual(5d, ExecuteFunction(new CountA(), 1d, 2m, 3, new DateTime(2012, 4, 1), "4", null, string.Empty));

	[TestMethod]
	public void CountAShouldIncludeEnumerableMembers() => Assert.AreEqual(3d, ExecuteFunction(new CountA(), 1d, FunctionsHelper.CreateArgs(12, 13)));

	[TestMethod]
	public void CountAShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
	{
		var result = ExecuteWithHiddenArg(new CountA(), a => a.ElementAt(0), 1d, FunctionsHelper.CreateArgs(12, 13));
		Assert.AreEqual(2d, result.Result);
	}

	[TestMethod]
	public void ProductShouldMultiplyArguments() => Assert.AreEqual(16d, ExecuteFunction(new Product(), 2d, 2d, 4d));

	[TestMethod]
	public void ProductShouldHandleEnumerable() => Assert.AreEqual(32d, ExecuteFunction(new Product(), 2d, 2d, FunctionsHelper.CreateArgs(4d, 2d)));

	[TestMethod]
	public void ProductShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
	{
		var result = ExecuteWithHiddenArg(new Product(), a => a.ElementAt(1), 2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
		Assert.AreEqual(16d, result.Result);
	}

	[TestMethod]
	public void ProductShouldHandleFirstItemIsEnumerable() => Assert.AreEqual(32d, ExecuteFunction(new Product(), FunctionsHelper.CreateArgs(4d, 2d), 2d, 2d));

	[TestMethod]
	public void VarShouldReturnCorrectResult()
	{
		var func = new Var();
		var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
		var result = func.Execute(args, _parsingContext);
		Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
	}

	[TestMethod]
	public void VarShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
	{
		var result = ExecuteWithHiddenArg(new Var(), a => a.Last(), 1, 2, 3, 4, 9);
		Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
	}

	[TestMethod]
	public void VarPShouldReturnCorrectResult() => Assert.AreEqual(1.25d, ExecuteFunction(new VarP(), 1, 2, 3, 4));

	[TestMethod]
	public void VarPShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
	{
		var result = ExecuteWithHiddenArg(new VarP(), a => a.Last(), 1, 2, 3, 4, 9);
		Assert.AreEqual(1.25d, result.Result);
	}
}

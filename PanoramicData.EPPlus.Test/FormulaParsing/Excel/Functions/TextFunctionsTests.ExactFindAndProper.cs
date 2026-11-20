using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class TextFunctionsTests
{
	[TestMethod]
	public void ExactShouldReturnTrueWhenTwoEqualStrings()
	{
		var func = new Exact();
		var result = func.Execute(FunctionsHelper.CreateArgs("abc", "abc"), _parsingContext);
		Assert.IsTrue((bool)result.Result);
	}

	[TestMethod]
	public void ExactShouldReturnTrueWhenEqualStringAndDouble()
	{
		var func = new Exact();
		var result = func.Execute(FunctionsHelper.CreateArgs("1", 1d), _parsingContext);
		Assert.IsTrue((bool)result.Result);
	}

	[TestMethod]
	public void ExactShouldReturnFalseWhenStringAndNull()
	{
		var func = new Exact();
		var result = func.Execute(FunctionsHelper.CreateArgs("1", null), _parsingContext);
		Assert.IsFalse((bool)result.Result);
	}

	[TestMethod]
	public void ExactShouldReturnFalseWhenTwoEqualStringsWithDifferentCase()
	{
		var func = new Exact();
		var result = func.Execute(FunctionsHelper.CreateArgs("abc", "Abc"), _parsingContext);
		Assert.IsFalse((bool)result.Result);
	}

	[TestMethod]
	public void FindShouldReturnIndexOfFoundPhrase()
	{
		var func = new Find();
		var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hej hopp"), _parsingContext);
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void FindShouldReturnIndexOfFoundPhraseBasedOnStartIndex()
	{
		var func = new Find();
		var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hopp hopp", 2), _parsingContext);
		Assert.AreEqual(6, result.Result);
	}

	[TestMethod]
	public void ProperShouldSetFirstLetterToUpperCase()
	{
		var func = new Proper();
		var result = func.Execute(FunctionsHelper.CreateArgs("this IS A tEst.wi3th SOME w0rds östEr"), _parsingContext);
		Assert.AreEqual("This Is A Test.Wi3Th Some W0Rds Öster", result.Result);
	}

	[TestMethod]
	public void HyperLinkShouldReturnArgIfOneArgIsSupplied()
	{
		var func = new Hyperlink();
		var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com"), _parsingContext);
		Assert.AreEqual("http://epplus.codeplex.com", result.Result);
	}

	[TestMethod]
	public void HyperLinkShouldReturnLastArgIfTwoArgsAreSupplied()
	{
		var func = new Hyperlink();
		var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com", "EPPlus"), _parsingContext);
		Assert.AreEqual("EPPlus", result.Result);
	}
}

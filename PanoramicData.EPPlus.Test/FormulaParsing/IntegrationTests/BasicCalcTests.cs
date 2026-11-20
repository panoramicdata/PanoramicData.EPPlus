using OfficeOpenXml.FormulaParsing;
using FakeItEasy;


namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests;

[TestClass]
public partial class BasicCalcTests : FormulaParserTestBase
{
	[TestInitialize]
	public void Setup()
	{
		var excelDataProvider = A.Fake<ExcelDataProvider>();
		_parser = new FormulaParser(excelDataProvider);
	}

	[TestMethod]
	public void ShouldAddIntegersCorrectly()
	{
		var result = _parser.Parse("1 + 2");
		Assert.AreEqual(3d, result);
	}

	[TestMethod]
	public void ShouldSubtractIntegersCorrectly()
	{
		var result = _parser.Parse("2 - 1");
		Assert.AreEqual(1d, result);
	}

	[TestMethod]
	public void ShouldMultiplyIntegersCorrectly()
	{
		var result = _parser.Parse("2 * 3");
		Assert.AreEqual(6d, result);
	}

	[TestMethod]
	public void ShouldDivideIntegersCorrectly()
	{
		var result = _parser.Parse("8 / 4");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void ShouldDivideDecimalWithIntegerCorrectly()
	{
		var result = _parser.Parse("2.5/2");
		Assert.AreEqual(1.25d, result);
	}

	[TestMethod]
	public void ShouldHandleExpCorrectly()
	{
		var result = _parser.Parse("2 ^ 4");
		Assert.AreEqual(16d, result);
	}

	[TestMethod]
	public void ShouldHandleExpWithDecimalCorrectly()
	{
		var result = _parser.Parse("2.5 ^ 2");
		Assert.AreEqual(6.25d, result);
	}

	[TestMethod]
	public void ShouldMultiplyDecimalWithDecimalCorrectly()
	{
		var result = _parser.Parse("2.5 * 1.5");
		Assert.AreEqual(3.75d, result);
	}
}

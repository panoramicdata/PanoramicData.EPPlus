using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public class ArgumentParserFactoryTests
{
	private ArgumentParserFactory _parserFactory;

	[TestInitialize]
	public void Setup() => _parserFactory = new ArgumentParserFactory();

	[TestMethod]
	public void ShouldReturnIntArgumentParserWhenDataTypeIsInteger()
	{
		var parser = _parserFactory.CreateArgumentParser(DataType.Integer);
		Assert.IsInstanceOfType<IntArgumentParser>(parser);
	}

	[TestMethod]
	public void ShouldReturnBoolArgumentParserWhenDataTypeIsBoolean()
	{
		var parser = _parserFactory.CreateArgumentParser(DataType.Boolean);
		Assert.IsInstanceOfType<BoolArgumentParser>(parser);
	}

	[TestMethod]
	public void ShouldReturnDoubleArgumentParserWhenDataTypeIsDecial()
	{
		var parser = _parserFactory.CreateArgumentParser(DataType.Decimal);
		Assert.IsInstanceOfType<DoubleArgumentParser>(parser);
	}
}

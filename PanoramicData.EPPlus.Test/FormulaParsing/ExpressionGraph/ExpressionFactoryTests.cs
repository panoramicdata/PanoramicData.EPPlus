using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExpressionGraph;

[TestClass]
public class ExpressionFactoryTests
{
	private ExpressionFactory _factory;
	private ParsingContext _parsingContext;

	[TestInitialize]
	public void Setup()
	{
		_parsingContext = ParsingContext.Create();
		var provider = A.Fake<ExcelDataProvider>();
		_factory = new ExpressionFactory(provider, _parsingContext);
	}

	[TestMethod]
	public void ShouldReturnIntegerExpressionWhenTokenIsInteger()
	{
		var token = new Token("2", TokenType.Integer);
		var expression = _factory.Create(token);
		Assert.IsInstanceOfType<IntegerExpression>(expression);
	}

	[TestMethod]
	public void ShouldReturnBooleanExpressionWhenTokenIsBoolean()
	{
		var token = new Token("true", TokenType.Boolean);
		var expression = _factory.Create(token);
		Assert.IsInstanceOfType<BooleanExpression>(expression);
	}

	[TestMethod]
	public void ShouldReturnDecimalExpressionWhenTokenIsDecimal()
	{
		var token = new Token("2.5", TokenType.Decimal);
		var expression = _factory.Create(token);
		Assert.IsInstanceOfType<DecimalExpression>(expression);
	}

	[TestMethod]
	public void ShouldReturnExcelRangeExpressionWhenTokenIsExcelAddress()
	{
		var token = new Token("A1", TokenType.ExcelAddress);
		var expression = _factory.Create(token);
		Assert.IsInstanceOfType<ExcelAddressExpression>(expression);
	}

	[TestMethod]
	public void ShouldReturnNamedValueExpressionWhenTokenIsNamedValue()
	{
		var token = new Token("NamedValue", TokenType.NameValue);
		var expression = _factory.Create(token);
		Assert.IsInstanceOfType<NamedValueExpression>(expression);
	}
}

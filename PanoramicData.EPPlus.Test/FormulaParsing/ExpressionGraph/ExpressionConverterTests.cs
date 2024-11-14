using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Globalization;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExpressionGraph;

[TestClass]
public class ExpressionConverterTests
{
	private ExpressionConverter _converter;

	[TestInitialize]
	public void Setup() => _converter = new ExpressionConverter();

	[TestMethod]
	public void ToStringExpressionShouldConvertIntegerExpressionToStringExpression()
	{
		var integerExpression = new IntegerExpression("2");
		var result = _converter.ToStringExpression(integerExpression);
		Assert.IsInstanceOfType<StringExpression>(result);
		Assert.AreEqual("2", result.Compile().Result);
	}

	[TestMethod]
	public void ToStringExpressionShouldCopyOperatorToStringExpression()
	{
		var integerExpression = new IntegerExpression("2")
		{
			Operator = Operator.Plus
		};
		var result = _converter.ToStringExpression(integerExpression);
		Assert.AreEqual(integerExpression.Operator, result.Operator);
	}

	[TestMethod]
	public void ToStringExpressionShouldConvertDecimalExpressionToStringExpression()
	{
		var decimalExpression = new DecimalExpression("2.5");
		var result = _converter.ToStringExpression(decimalExpression);
		Assert.IsInstanceOfType<StringExpression>(result);
		Assert.AreEqual($"2{CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator}5", result.Compile().Result);
	}

	[TestMethod]
	public void FromCompileResultShouldCreateIntegerExpressionIfCompileResultIsInteger()
	{
		var compileResult = new CompileResult(1, DataType.Integer);
		var result = _converter.FromCompileResult(compileResult);
		Assert.IsInstanceOfType<IntegerExpression>(result);
		Assert.AreEqual(1d, result.Compile().Result);
	}

	[TestMethod]
	public void FromCompileResultShouldCreateStringExpressionIfCompileResultIsString()
	{
		var compileResult = new CompileResult("abc", DataType.String);
		var result = _converter.FromCompileResult(compileResult);
		Assert.IsInstanceOfType<StringExpression>(result);
		Assert.AreEqual("abc", result.Compile().Result);
	}

	[TestMethod]
	public void FromCompileResultShouldCreateDecimalExpressionIfCompileResultIsDecimal()
	{
		var compileResult = new CompileResult(2.5d, DataType.Decimal);
		var result = _converter.FromCompileResult(compileResult);
		Assert.IsInstanceOfType<DecimalExpression>(result);
		Assert.AreEqual(2.5d, result.Compile().Result);
	}

	[TestMethod]
	public void FromCompileResultShouldCreateBooleanExpressionIfCompileResultIsBoolean()
	{
		var compileResult = new CompileResult("true", DataType.Boolean);
		var result = _converter.FromCompileResult(compileResult);
		Assert.IsInstanceOfType<BooleanExpression>(result);
		Assert.IsTrue((bool)result.Compile().Result);
	}
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class RoundTests
{
	[TestMethod]
	public void RoundPositiveToOnesDownLiteral()
	{
		Round round = new();
		var value1 = 123.45;
		var digits = 0;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(123D, result.Result);
	}
	[TestMethod]
	public void RoundPositiveToOnesUpLiteral()
	{
		Round round = new();
		var value1 = 123.65;
		var digits = 0;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(124D, result.Result);
	}

	[TestMethod]
	public void RoundPositiveToTenthsDownLiteral()
	{
		Round round = new();
		var value1 = 123.44;
		var digits = 1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(123.4D, result.Result);
	}
	[TestMethod]
	public void RoundPositiveToTenthsUpLiteral()
	{
		Round round = new();
		var value1 = 123.456;
		var digits = 1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(123.5D, result.Result);
	}
	[TestMethod]
	public void RoundPositiveToTensDownLiteral()
	{
		Round round = new();
		double value1 = 124;
		var digits = -1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(120D, result.Result);
	}
	[TestMethod]
	public void RoundPositiveToTensUpLiteral()
	{
		Round round = new();
		double value1 = 125;
		var digits = -1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(130D, result.Result);
	}

	[TestMethod]
	public void RoundNegativeToTensDownLiteral()
	{
		Round round = new();
		double value1 = -124;
		var digits = -1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(-120D, result.Result);
	}
	[TestMethod]
	public void RoundNegativeToTensUpLiteral()
	{
		Round round = new();
		double value1 = -125;
		var digits = -1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(-130D, result.Result);
	}
	[TestMethod]
	public void RoundNegativeToTenthsDownLiteral()
	{
		Round round = new();
		var value1 = -123.44;
		var digits = 1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(-123.4D, result.Result);
	}
	[TestMethod]
	public void RoundNegativeToTenthsUpLiteral()
	{
		Round round = new();
		var value1 = -123.456;
		var digits = 1;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(-123.5D, result.Result);
	}
	[TestMethod]
	public void RoundNegativeMidwayLiteral()
	{
		Round round = new();
		var value1 = -123.5;
		var digits = 0;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(-124D, result.Result);
	}
	[TestMethod]
	public void RoundPositiveMidwayLiteral()
	{
		Round round = new();
		var value1 = 123.5;
		var digits = 0;
		var result = round.Execute(new FunctionArgument[]
		{
			new(value1),
			new(digits)
		}, ParsingContext.Create());
		Assert.AreEqual(124D, result.Result);
	}
}

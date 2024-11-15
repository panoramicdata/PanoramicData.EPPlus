﻿using FakeItEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExpressionGraph;

[TestClass]
public class ExpressionGraphBuilderTests
{
	private ExpressionGraphBuilder _graphBuilder;
	private ExcelDataProvider _excelDataProvider;

	[TestInitialize]
	public void Setup()
	{
		_excelDataProvider = A.Fake<ExcelDataProvider>();
		var parsingContext = ParsingContext.Create();
		_graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, parsingContext);
	}

	[TestCleanup]
	public void Cleanup()
	{

	}

	[TestMethod]
	public void BuildShouldNotUseStringIdentifyersWhenBuildingStringExpression()
	{
		var tokens = new List<Token>
		{
			new("'", TokenType.String),
			new("abc", TokenType.StringContent),
			new("'", TokenType.String)
		};

		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(1, result.Expressions.Count());
	}

	[TestMethod]
	public void BuildShouldNotEvaluateExpressionsWithinAString()
	{
		var tokens = new List<Token>
		{
			new("'", TokenType.String),
			new("1 + 2", TokenType.StringContent),
			new("'", TokenType.String)
		};

		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual("1 + 2", result.Expressions.First().Compile().Result);
	}

	[TestMethod]
	public void BuildShouldSetOperatorOnGroupExpressionCorrectly()
	{
		var tokens = new List<Token>
		{
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new("+", TokenType.Operator),
			new("4", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis),
			new("*", TokenType.Operator),
			new("2", TokenType.Integer)
		};
		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(Operator.Multiply.Operator, result.Expressions.First().Operator.Operator);

	}

	[TestMethod]
	public void BuildShouldSetChildrenOnGroupExpression()
	{
		var tokens = new List<Token>
		{
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new("+", TokenType.Operator),
			new("4", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis),
			new("*", TokenType.Operator),
			new("2", TokenType.Integer)
		};
		var result = _graphBuilder.Build(tokens);

		Assert.IsInstanceOfType<GroupExpression>(result.Expressions.First());
		Assert.AreEqual(2, result.Expressions.First().Children.Count());
	}

	[TestMethod]
	public void BuildShouldSetNextOnGroupedExpression()
	{
		var tokens = new List<Token>
		{
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new("+", TokenType.Operator),
			new("4", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis),
			new("*", TokenType.Operator),
			new("2", TokenType.Integer)
		};
		var result = _graphBuilder.Build(tokens);

		Assert.IsNotNull(result.Expressions.First().Next);
		Assert.IsInstanceOfType<IntegerExpression>(result.Expressions.First().Next);

	}

	[TestMethod]
	public void BuildShouldBuildFunctionExpressionIfFirstTokenIsFunction()
	{
		var tokens = new List<Token>
		{
			new("CStr", TokenType.Function),
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis),
		};
		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(1, result.Expressions.Count());
		Assert.IsInstanceOfType<FunctionExpression>(result.Expressions.First());
	}

	[TestMethod]
	public void BuildShouldSetChildrenOnFunctionExpression()
	{
		var tokens = new List<Token>
		{
			new("CStr", TokenType.Function),
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis)
		};
		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(1, result.Expressions.First().Children.Count());
		Assert.IsInstanceOfType<GroupExpression>(result.Expressions.First().Children.First());
		Assert.IsInstanceOfType<IntegerExpression>(result.Expressions.First().Children.First().Children.First());
		Assert.AreEqual(2d, result.Expressions.First().Children.First().Compile().Result);
	}

	[TestMethod]
	public void BuildShouldAddOperatorToFunctionExpression()
	{
		var tokens = new List<Token>
		{
			new("CStr", TokenType.Function),
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis),
			new("&", TokenType.Operator),
			new("A", TokenType.StringContent)
		};
		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(1, result.Expressions.First().Children.Count());
		Assert.AreEqual(2, result.Expressions.Count());
	}

	[TestMethod]
	public void BuildShouldAddCommaSeparatedFunctionArgumentsAsChildrenToFunctionExpression()
	{
		var tokens = new List<Token>
		{
			new("Text", TokenType.Function),
			new("(", TokenType.OpeningParenthesis),
			new("2", TokenType.Integer),
			new(",", TokenType.Comma),
			new("3", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis),
			new("&", TokenType.Operator),
			new("A", TokenType.StringContent)
		};

		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(2, result.Expressions.First().Children.Count());
	}

	[TestMethod]
	public void BuildShouldCreateASingleExpressionOutOfANegatorAndANumericToken()
	{
		var tokens = new List<Token>
		{
			new("-", TokenType.Negator),
			new("2", TokenType.Integer),
		};

		var result = _graphBuilder.Build(tokens);

		Assert.AreEqual(1, result.Expressions.Count());
		Assert.AreEqual(-2d, result.Expressions.First().Compile().Result);
	}

	[TestMethod]
	public void BuildShouldHandleEnumerableTokens()
	{
		var tokens = new List<Token>
		{
			new("Text", TokenType.Function),
			new("(", TokenType.OpeningParenthesis),
			new("{", TokenType.OpeningEnumerable),
			new("2", TokenType.Integer),
			new(",", TokenType.Comma),
			new("3", TokenType.Integer),
			new("}", TokenType.ClosingEnumerable),
			new(")", TokenType.ClosingParenthesis)
		};

		var result = _graphBuilder.Build(tokens);
		var funcArgExpression = result.Expressions.First().Children.First();
		Assert.IsInstanceOfType<FunctionArgumentExpression>(funcArgExpression);

		var enumerableExpression = funcArgExpression.Children.First();

		Assert.IsInstanceOfType<EnumerableExpression>(enumerableExpression);
		Assert.AreEqual(2, enumerableExpression.Children.Count(), "Enumerable.Count was not 2");
	}

	[TestMethod]
	public void ShouldHandleInnerFunctionCall2()
	{
		var ctx = ParsingContext.Create();
		const string formula = "IF(3>2;\"Yes\";\"No\")";
		var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
		var tokens = tokenizer.Tokenize(formula);
		var expression = _graphBuilder.Build(tokens);
		Assert.AreEqual(1, expression.Expressions.Count());

		var compiler = new ExpressionCompiler(new ExpressionConverter(), new CompileStrategyFactory());
		var result = compiler.Compile(expression.Expressions);
		Assert.AreEqual("Yes", result.Result);
	}

	[TestMethod]
	public void ShouldHandleInnerFunctionCall3()
	{
		var ctx = ParsingContext.Create();
		const string formula = "IF(I10>=0;IF(O10>I10;((O10-I10)*$B10)/$C$27;IF(O10<0;(O10*$B10)/$C$27;\"\"));IF(O10<0;((O10-I10)*$B10)/$C$27;IF(O10>0;(O10*$B10)/$C$27;)))";
		var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
		var tokens = tokenizer.Tokenize(formula);
		var expression = _graphBuilder.Build(tokens);
		Assert.AreEqual(1, expression.Expressions.Count());
		var exp1 = expression.Expressions.First();
		Assert.AreEqual(3, exp1.Children.Count());
	}
	[TestMethod]
	public void RemoveDuplicateOperators1()
	{
		var ctx = ParsingContext.Create();
		const string formula = "++1--2++-3+-1----3-+2";
		var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
		var tokens = tokenizer.Tokenize(formula).ToList();
		var expression = _graphBuilder.Build(tokens);
		Assert.AreEqual(11, tokens.Count);
		Assert.AreEqual("+", tokens[1].Value);
		Assert.AreEqual("-", tokens[3].Value);
		Assert.AreEqual("-", tokens[5].Value);
		Assert.AreEqual("+", tokens[7].Value);
		Assert.AreEqual("-", tokens[9].Value);
	}
	[TestMethod]
	public void RemoveDuplicateOperators2()
	{
		var ctx = ParsingContext.Create();
		const string formula = "++-1--(---2)++-3+-1----3-+2";
		var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
		var tokens = tokenizer.Tokenize(formula).ToList();
	}

	[TestMethod]
	public void BuildExcelAddressExpressionSimple()
	{
		var tokens = new List<Token>
		{
			new("A1", TokenType.ExcelAddress)
		};

		var result = _graphBuilder.Build(tokens);
		Assert.IsInstanceOfType<ExcelAddressExpression>(result.Expressions.First());
	}
}

﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Linq;
using ExpGraph = OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExpressionGraph;

[TestClass]
public class ExpressionCompilerTests
{
	private ExpressionCompiler _expressionCompiler;
	private ExpGraph _graph;

	[TestInitialize]
	public void Setup()
	{
		_expressionCompiler = new ExpressionCompiler();
		_graph = new ExpGraph();
	}

	[TestMethod]
	public void ShouldCompileTwoInterExpressionsToCorrectResult()
	{
		var exp1 = new IntegerExpression("2")
		{
			Operator = Operator.Plus
		};
		_graph.Add(exp1);
		var exp2 = new IntegerExpression("2");
		_graph.Add(exp2);

		var result = _expressionCompiler.Compile(_graph.Expressions);

		Assert.AreEqual(4d, result.Result);
	}


	[TestMethod]
	public void CompileShouldMultiplyGroupExpressionWithFollowingIntegerExpression()
	{
		var groupExpression = new GroupExpression(false);
		groupExpression.AddChild(new IntegerExpression("2"));
		groupExpression.Children.First().Operator = Operator.Plus;
		groupExpression.AddChild(new IntegerExpression("3"));
		groupExpression.Operator = Operator.Multiply;

		_graph.Add(groupExpression);
		_graph.Add(new IntegerExpression("2"));

		var result = _expressionCompiler.Compile(_graph.Expressions);

		Assert.AreEqual(10d, result.Result);
	}

	[TestMethod]
	public void CompileShouldCalculateMultipleExpressionsAccordingToPrecedence()
	{
		var exp1 = new IntegerExpression("2")
		{
			Operator = Operator.Multiply
		};
		_graph.Add(exp1);
		var exp2 = new IntegerExpression("2")
		{
			Operator = Operator.Plus
		};
		_graph.Add(exp2);
		var exp3 = new IntegerExpression("2")
		{
			Operator = Operator.Multiply
		};
		_graph.Add(exp3);
		var exp4 = new IntegerExpression("2");
		_graph.Add(exp4);

		var result = _expressionCompiler.Compile(_graph.Expressions);

		Assert.AreEqual(8d, result.Result);
	}
}

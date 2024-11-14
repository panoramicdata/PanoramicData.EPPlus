using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;

namespace PanoramicData.EPPlus.Test.FormulaParsing.LexicalAnalysis;

[TestClass]
public class SyntacticAnalyzerTests
{
	private SyntacticAnalyzer _analyser;

	[TestInitialize]
	public void Setup() => _analyser = new SyntacticAnalyzer();

	[TestMethod]
	public void ShouldPassIfParenthesisAreWellformed()
	{
		var input = new List<Token>
		{
			new("(", TokenType.OpeningParenthesis),
			new("1", TokenType.Integer),
			new("+", TokenType.Operator),
			new("2", TokenType.Integer),
			new(")", TokenType.ClosingParenthesis)
		};
		_analyser.Analyze(input);
	}

	[TestMethod, ExpectedException(typeof(FormatException))]
	public void ShouldThrowExceptionIfParenthesesAreNotWellformed()
	{
		var input = new List<Token>
		{
			new("(", TokenType.OpeningParenthesis),
			new("1", TokenType.Integer),
			new("+", TokenType.Operator),
			new("2", TokenType.Integer)
		};
		_analyser.Analyze(input);
	}

	[TestMethod]
	public void ShouldPassIfStringIsWellformed()
	{
		var input = new List<Token>
		{
			new("'", TokenType.String),
			new("abc123", TokenType.StringContent),
			new("'", TokenType.String)
		};
		_analyser.Analyze(input);
	}

	[TestMethod, ExpectedException(typeof(FormatException))]
	public void ShouldThrowExceptionIfStringHasNotClosing()
	{
		var input = new List<Token>
		{
			new("'", TokenType.String),
			new("abc123", TokenType.StringContent)
		};
		_analyser.Analyze(input);
	}


	[TestMethod, ExpectedException(typeof(UnrecognizedTokenException))]
	public void ShouldThrowExceptionIfThereIsAnUnrecognizedToken()
	{
		var input = new List<Token>
		{
			new("abc123", TokenType.Unrecognized)
		};
		_analyser.Analyze(input);
	}
}

using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace PanoramicData.EPPlus.Test.FormulaParsing.LexicalAnalysis;

public partial class SourceCodeTokenizerTests
{
	[TestMethod]
	public void TokenizeStripsLeadingPlusSign()
	{
		AssertTokenTypes(@"+3-3", TokenType.Integer, TokenType.Operator, TokenType.Integer);
	}

	[TestMethod]
	public void TokenizeStripsLeadingDoubleNegator()
	{
		AssertTokenTypes(@"--3-3", TokenType.Integer, TokenType.Operator, TokenType.Integer);
	}

	[TestMethod]
	public void TokenizeHandlesPositiveNegator()
	{
		AssertTokenTypes(@"+-3-3", TokenType.Negator, TokenType.Integer, TokenType.Operator, TokenType.Integer);
	}

	[TestMethod]
	public void TokenizeHandlesNegatorPositive()
	{
		AssertTokenTypes(@"-+3-3", TokenType.Negator, TokenType.Integer, TokenType.Operator, TokenType.Integer);
	}

	[TestMethod]
	public void TokenizeStripsLeadingPlusSignFromFirstFunctionArgument()
	{
		AssertTokenTypes(@"SUM(+3-3,5)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.Comma, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeStripsLeadingPlusSignFromSecondFunctionArgument()
	{
		AssertTokenTypes(@"SUM(5,+3-3)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Integer, TokenType.Comma, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeStripsLeadingDoubleNegatorFromFirstFunctionArgument()
	{
		AssertTokenTypes(@"SUM(--3-3,5)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.Comma, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeStripsLeadingDoubleNegatorFromSecondFunctionArgument()
	{
		AssertTokenTypes(@"SUM(5,--3-3)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Integer, TokenType.Comma, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeHandlesPositiveNegatorAsFirstFunctionArgument()
	{
		AssertTokenTypes(@"SUM(+-3-3,5)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Negator, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.Comma, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeHandlesNegatorPositiveAsFirstFunctionArgument()
	{
		AssertTokenTypes(@"SUM(-+3-3,5)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Negator, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.Comma, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeHandlesPositiveNegatorAsSecondFunctionArgument()
	{
		AssertTokenTypes(@"SUM(5,+-3-3)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Integer, TokenType.Comma, TokenType.Negator, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.ClosingParenthesis);
	}

	[TestMethod]
	public void TokenizeHandlesNegatorPositiveAsSecondFunctionArgument()
	{
		AssertTokenTypes(@"SUM(5,-+3-3)", TokenType.Function, TokenType.OpeningParenthesis, TokenType.Integer, TokenType.Comma, TokenType.Negator, TokenType.Integer, TokenType.Operator, TokenType.Integer, TokenType.ClosingParenthesis);
	}
}

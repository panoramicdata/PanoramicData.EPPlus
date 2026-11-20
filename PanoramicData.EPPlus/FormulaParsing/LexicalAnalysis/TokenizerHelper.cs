namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

internal static class TokenizerHelper
{
	internal static bool PrecedesNegator(Token t) => t == null
		|| t.TokenType == TokenType.Operator
		|| t.TokenType == TokenType.OpeningParenthesis
		|| t.TokenType == TokenType.Comma
		|| t.TokenType == TokenType.SemiColon
		|| t.TokenType == TokenType.OpeningEnumerable;

	internal static Token CreateCurrentToken(ITokenFactory tokenFactory, TokenizerContext context, string worksheet)
	{
		if (context.CurrentToken == "-" && context.LastToken == null && context.LastToken.TokenType == TokenType.Operator)
		{
			return new Token("-", TokenType.Negator);
		}

		return tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
	}
}

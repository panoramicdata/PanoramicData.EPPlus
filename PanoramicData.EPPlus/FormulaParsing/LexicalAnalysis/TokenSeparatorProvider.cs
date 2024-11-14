/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class TokenSeparatorProvider : ITokenSeparatorProvider
{
	private static readonly Dictionary<string, Token> _tokens;

	static TokenSeparatorProvider()
	{
		_tokens = new Dictionary<string, Token>
		{
			{ "+", new Token("+", TokenType.Operator) },
			{ "-", new Token("-", TokenType.Operator) },
			{ "*", new Token("*", TokenType.Operator) },
			{ "/", new Token("/", TokenType.Operator) },
			{ "^", new Token("^", TokenType.Operator) },
			{ "&", new Token("&", TokenType.Operator) },
			{ ">", new Token(">", TokenType.Operator) },
			{ "<", new Token("<", TokenType.Operator) },
			{ "=", new Token("=", TokenType.Operator) },
			{ "<=", new Token("<=", TokenType.Operator) },
			{ ">=", new Token(">=", TokenType.Operator) },
			{ "<>", new Token("<>", TokenType.Operator) },
			{ "(", new Token("(", TokenType.OpeningParenthesis) },
			{ ")", new Token(")", TokenType.ClosingParenthesis) },
			{ "{", new Token("{", TokenType.OpeningEnumerable) },
			{ "}", new Token("}", TokenType.ClosingEnumerable) },
			{ "'", new Token("'", TokenType.WorksheetName) },
			{ "\"", new Token("\"", TokenType.String) },
			{ ",", new Token(",", TokenType.Comma) },
			{ ";", new Token(";", TokenType.SemiColon) },
			{ "[", new Token("[", TokenType.OpeningBracket) },
			{ "]", new Token("]", TokenType.ClosingBracket) },
			{ "%", new Token("%", TokenType.Percent) }
		};
	}

	IDictionary<string, Token> ITokenSeparatorProvider.Tokens => _tokens;

	public bool IsOperator(string item)
	{
		if (_tokens.TryGetValue(item, out var token))
		{
			if (token.TokenType == TokenType.Operator)
			{
				return true;
			}
		}

		return false;
	}

	public bool IsPossibleLastPartOfMultipleCharOperator(string part) => part is "=" or ">";
}

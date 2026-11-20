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
 * Mats Alm   		                Added       		        2015-12-28
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public class TokenHandler : ITokenIndexProvider
{
	public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
	{
		_context = context;
		_tokenFactory = tokenFactory;
		_tokenProvider = tokenProvider;
	}

	private readonly TokenizerContext _context;
	private readonly ITokenSeparatorProvider _tokenProvider;
	private readonly ITokenFactory _tokenFactory;
	private int _tokenIndex = -1;

	public string Worksheet { get; set; }

	public bool HasMore() => _tokenIndex < _context.FormulaChars.Length - 1;

	public void Next()
	{
		_tokenIndex++;
		Handle();
	}

	private void Handle()
	{
		var c = _context.FormulaChars[_tokenIndex];
		if (CharIsTokenSeparator(c, out var tokenSeparator))
		{
			HandleTokenSeparator(c, tokenSeparator);
			return;
		}

		_context.AppendToCurrentToken(c);
	}

	private void HandleTokenSeparator(char c, Token tokenSeparator)
	{
		if (TokenSeparatorHandler.Handle(c, tokenSeparator, _context, this))
		{
			return;
		}

		if (_context.CurrentTokenHasValue)
		{
			AddCurrentToken(tokenSeparator);
		}

		if (tokenSeparator.Value == "-" && TokenIsNegator(_context))
		{
			_context.AddToken(new Token("-", TokenType.Negator));
			return;
		}

		_context.AddToken(tokenSeparator);
		_context.NewToken();
	}

	private void AddCurrentToken(Token tokenSeparator)
	{
		if (Regex.IsMatch(_context.CurrentToken, "^\"*$"))
		{
			_context.AddToken(_tokenFactory.Create(_context.CurrentToken, TokenType.StringContent));
		}
		else
		{
			_context.AddToken(CreateToken(_context, Worksheet));
		}

		// An address or name followed by an opening parenthesis is a function.
		if (tokenSeparator.TokenType == TokenType.OpeningParenthesis && (_context.LastToken.TokenType == TokenType.ExcelAddress || _context.LastToken.TokenType == TokenType.NameValue))
		{
			_context.LastToken.TokenType = TokenType.Function;
		}
	}

	private bool CharIsTokenSeparator(char c, out Token token)
	{
		var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
		token = result ? _tokenProvider.Tokens[c.ToString()] : null;
		return result;
	}

	private static bool TokenIsNegator(TokenizerContext context) => TokenizerHelper.PrecedesNegator(context.LastToken);

	private Token CreateToken(TokenizerContext context, string worksheet) => TokenizerHelper.CreateCurrentToken(_tokenFactory, context, worksheet);

	int ITokenIndexProvider.Index => _tokenIndex;


	void ITokenIndexProvider.MoveIndexPointerForward() => _tokenIndex++;
}

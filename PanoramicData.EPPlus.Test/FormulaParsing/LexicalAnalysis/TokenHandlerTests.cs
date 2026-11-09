using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.LexicalAnalysis;

[TestClass]
public class TokenHandlerTests
{
	private TokenHandler _handler;

	[TestInitialize]
	public void Init()
	{
		var tokenizerContext = new TokenizerContext("test");
		InitHandler(tokenizerContext);
	}

	private void InitHandler(TokenizerContext context)
	{
		var parsingContext = ParsingContext.Create();
		var tokenFactory = new TokenFactory(parsingContext.Configuration.FunctionRepository, null);
		_handler = new TokenHandler(context, tokenFactory, new TokenSeparatorProvider());
	}

	[TestMethod]
	public void HasMoreTokensShouldBeTrueWhenTokensExists() => Assert.IsTrue(_handler.HasMore());

	[TestMethod]
	public void HasMoreTokensShouldBeFalseWhenAllAreHandled()
	{
		for (var x = 0; x < "test".Length; x++)
		{
			_handler.Next();
		}

		Assert.IsFalse(_handler.HasMore());
	}
}

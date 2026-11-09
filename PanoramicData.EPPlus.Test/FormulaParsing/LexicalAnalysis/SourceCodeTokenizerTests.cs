using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.LexicalAnalysis;

[TestClass]
public partial class SourceCodeTokenizerTests
{
	private SourceCodeTokenizer _tokenizer;

	[TestInitialize]
	public void Setup()
	{
		var context = ParsingContext.Create();
		_tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
	}

	private void AssertTokenTypes(string input, params TokenType[] expected)
	{
		var tokens = _tokenizer.Tokenize(input);
		Assert.AreEqual(expected.Length, tokens.Count());
		for (var i = 0; i < expected.Length; i++)
		{
			Assert.AreEqual(expected[i], tokens.ElementAt(i).TokenType);
		}
	}
}

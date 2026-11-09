using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using FakeItEasy;

namespace PanoramicData.EPPlus.Test.FormulaParsing.ExcelUtilities;

[TestClass]
public class CellReferenceProviderTests
{
	private ExcelDataProvider _provider;

	[TestInitialize]
	public void Setup()
	{
		_provider = A.Fake<ExcelDataProvider>();
		A.CallTo(() => _provider.ExcelMaxRows).Returns(5000);
	}

	private IEnumerable<string> GetReferencedAddresses(string formula)
	{
		var parsingContext = ParsingContext.Create();
		parsingContext.Scopes.NewScope(RangeAddress.Empty);
		parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
		parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider);
		var provider = new CellReferenceProvider();
		return provider.GetReferencedAddresses(formula, parsingContext);
	}

	[TestMethod]
	public void ShouldReturnReferencedSingleAddress()
	{
		var result = GetReferencedAddresses("A1");
		Assert.AreEqual("A1", result.First());
	}

	[TestMethod]
	public void ShouldReturnReferencedMultipleAddresses()
	{
		var result = GetReferencedAddresses("A1:A2");
		Assert.AreEqual("A1", result.First());
		Assert.AreEqual("A2", result.Last());
	}
}

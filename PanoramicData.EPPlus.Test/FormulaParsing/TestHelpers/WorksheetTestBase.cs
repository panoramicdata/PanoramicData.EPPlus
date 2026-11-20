using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

/// <summary>Per-test parsing context and an in-memory package with one worksheet named "test".</summary>
public abstract class WorksheetTestBase
{
	private ExcelPackage _package;

	private protected ParsingContext Context { get; private set; }

	private protected ExcelWorksheet Worksheet { get; private set; }

	[TestInitialize]
	public void Initialize()
	{
		Context = ParsingContext.Create();
		_package = new ExcelPackage(new MemoryStream());
		Worksheet = _package.Workbook.Worksheets.Add("test");
	}

	[TestCleanup]
	public void Cleanup() => _package.Dispose();
}

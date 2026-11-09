using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public partial class MathFunctionsTests
{
	private ParsingContext _parsingContext;

	[TestInitialize]
	public void Initialize()
	{
		_parsingContext = ParsingContext.Create();
		_parsingContext.Scopes.NewScope(RangeAddress.Empty);
	}

	private object ExecuteFunction(ExcelFunction func, params object[] args)
		=> func.Execute(FunctionsHelper.CreateArgs(args), _parsingContext).Result;

	private double ExecuteFunctionRounded(ExcelFunction func, int digits, params object[] args)
		=> System.Math.Round((double)ExecuteFunction(func, args), digits);

	private CompileResult ExecuteWithHiddenArg(HiddenValuesHandlingFunction func, Func<IEnumerable<FunctionArgument>, FunctionArgument> hiddenArg, params object[] args)
	{
		func.IgnoreHiddenValues = true;
		var arguments = FunctionsHelper.CreateArgs(args);
		hiddenArg(arguments).SetExcelStateFlag(ExcelCellState.HiddenCell);
		return func.Execute(arguments, _parsingContext);
	}
}

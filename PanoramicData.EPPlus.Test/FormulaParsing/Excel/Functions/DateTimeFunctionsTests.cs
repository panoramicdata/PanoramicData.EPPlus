using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

[TestClass]
public partial class DateTimeFunctionsTests
{
	private readonly ParsingContext _parsingContext = ParsingContext.Create();

	private static double GetTime(int hour, int minute, int second)
	{
		var secInADay = DateTime.Today.AddDays(1).Subtract(DateTime.Today).TotalSeconds;
		var secondsOfExample = (double)(hour * 60 * 60 + minute * 60 + second);
		return secondsOfExample / secInADay;
	}

	private static object CalculateA1(string formula)
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("test");
		ws.Cells["A1"].Formula = formula;
		ws.Calculate();
		return ws.Cells["A1"].Value;
	}

	private object ExecuteWorkday(DateTime inputDate, int days)
		=> new Workday().Execute(FunctionsHelper.CreateArgs(inputDate.ToOADate(), days), _parsingContext).Result;

	private CompileResult ExecuteOnDates(ExcelFunction func, DateTime date1, DateTime date2, params object[] extraArgs)
	{
		object[] args = [date1.ToOADate(), date2.ToOADate(), .. extraArgs];
		return func.Execute(FunctionsHelper.CreateArgs(args), _parsingContext);
	}
}

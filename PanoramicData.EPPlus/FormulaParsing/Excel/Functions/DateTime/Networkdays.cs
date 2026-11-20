using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Networkdays : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
		ValidateArguments(functionArguments, 2);
		var startDate = ArgToOADate(functionArguments, 0);
		var endDate = ArgToOADate(functionArguments, 1);
		var result = new WorkdayCalculator().CalculateNumberOfWorkdays(startDate, endDate, functionArguments, 2);
		return new CompileResult(result.NumberOfWorkdays, DataType.Integer);
	}
}

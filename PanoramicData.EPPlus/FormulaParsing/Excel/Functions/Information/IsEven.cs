using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsEven : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		=> CreateResult(GetFirstArgumentAsFlooredInteger(arguments) % 2 == 0, DataType.Boolean);
}

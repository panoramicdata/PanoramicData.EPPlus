using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Sumsq : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		=> SumOverArguments(arguments, arg => Calculate(arg, context));

	private double Calculate(FunctionArgument arg, ParsingContext context, bool isInArray = false)
	{
		var retVal = 0d;
		if (ShouldIgnore(arg))
		{
			return retVal;
		}

		if (arg.Value is IEnumerable<FunctionArgument> items)
		{
			foreach (var item in items)
			{
				retVal += Calculate(item, context, true);
			}
		}
		else if (arg.Value is ExcelDataProvider.IRangeInfo cs)
		{
			retVal += SumVisibleCells(cs, context, value => System.Math.Pow(value, 2));
		}
		else
		{
			CheckForAndHandleExcelError(arg);
			if (IsNumericString(arg.Value) && !isInArray)
			{
				var value = ConvertUtil.GetValueDouble(arg.Value);
				return System.Math.Pow(value, 2);
			}

			var ignoreBool = isInArray;
			retVal += System.Math.Pow(ConvertUtil.GetValueDouble(arg.Value, ignoreBool), 2);
		}

		return retVal;
	}
}

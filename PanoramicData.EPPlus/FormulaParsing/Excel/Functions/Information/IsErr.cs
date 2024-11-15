﻿using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsErr : ErrorHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		var isError = new IsError();
		var result = isError.Execute(arguments, context);
		if ((bool)result.Result)
		{
			var arg = GetFirstValue(arguments);
			if (arg is ExcelDataProvider.IRangeInfo r)
			{
				if (r.GetValue(r.Address._fromRow, r.Address._fromCol) is ExcelErrorValue e && e.Type == eErrorType.NA)
				{
					return CreateResult(false, DataType.Boolean);
				}
			}
			else
			{
				if (arg is ExcelErrorValue && ((ExcelErrorValue)arg).Type == eErrorType.NA)
				{
					return CreateResult(false, DataType.Boolean);
				}
			}
		}

		return result;
	}

	public override CompileResult HandleError(string errorCode) => CreateResult(true, DataType.Boolean);
}

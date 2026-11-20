using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IfNaFunctionCompiler : FunctionCompiler
{
	public IfNaFunctionCompiler(ExcelFunction function, ParsingContext context)
		: base(function, context)
	{
	}

	public override CompileResult Compile(IEnumerable<Expression> children)
	{
		if (children.Count() != 2)
		{
			return new CompileResult(eErrorType.Value);
		}

		return CompileWithFallback(children, result => result.DataType == DataType.ExcelError && Equals(result.Result, ExcelErrorValue.Create(eErrorType.NA)));
	}
}

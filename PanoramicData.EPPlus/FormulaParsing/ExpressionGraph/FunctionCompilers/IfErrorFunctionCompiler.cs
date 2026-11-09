using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IfErrorFunctionCompiler : FunctionCompiler
{
	public IfErrorFunctionCompiler(ExcelFunction function, ParsingContext context)
		: base(function, context)
	{
		Require.That(function).Named("function").IsNotNull();
	}

	public override CompileResult Compile(IEnumerable<Expression> children)
	{
		if (children.Count() != 2)
		{
			throw new ExcelErrorValueException(eErrorType.Value);
		}

		return CompileWithFallback(children, result => result.DataType == DataType.ExcelError);
	}
}

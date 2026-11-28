using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Rand : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		Span<byte> bytes = stackalloc byte[8];
		RandomNumberGenerator.Fill(bytes);
		// The top 53 bits give a uniform double in [0, 1), as Excel's RAND() returns.
		var val = (BitConverter.ToUInt64(bytes) >> 11) * (1.0 / (1UL << 53));
		return CreateResult(val, DataType.Decimal);
	}
}

using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

/// <summary>
/// Simple implementation of DateValue function, just using .NET built-in
/// function System.DateTime.TryParse, based on current culture
/// </summary>
public class DateValue : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		var dateString = ArgToString(arguments, 0);
		return Execute(dateString);
	}

	internal CompileResult Execute(string dateString) => CreateDateResultFromString(dateString, result => result.ToOADate());
}

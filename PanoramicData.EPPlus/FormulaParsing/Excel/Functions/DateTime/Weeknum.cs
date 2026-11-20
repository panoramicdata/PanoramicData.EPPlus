using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Weeknum : ExcelFunction
{
	public override ExpressionGraph.CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1, eErrorType.Value);
		var dateSerial = ArgToDecimal(arguments, 0);
		var date = System.DateTime.FromOADate(dateSerial);
		var startDay = DayOfWeek.Sunday;
		if (arguments.Count() > 1)
		{
			startDay = GetStartDay(ArgToInt(arguments, 1));
		}

		if (DateTimeFormatInfo.CurrentInfo == null)
		{
			throw new InvalidOperationException(
				"Could not execute Weeknum function because DateTimeFormatInfo.CurrentInfo was null");
		}

		var week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay,
																		 startDay);
		return CreateResult(week, DataType.Integer);
	}

	private static DayOfWeek GetStartDay(int argStartDay) => argStartDay switch
	{
		1 => DayOfWeek.Sunday,
		2 or 11 => DayOfWeek.Monday,
		12 => DayOfWeek.Tuesday,
		13 => DayOfWeek.Wednesday,
		14 => DayOfWeek.Thursday,
		15 => DayOfWeek.Friday,
		16 => DayOfWeek.Saturday,
		_ => throw new Exceptions.ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(eErrorType.Num))
	};
}

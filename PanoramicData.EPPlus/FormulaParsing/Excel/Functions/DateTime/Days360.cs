using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Days360 : ExcelFunction
{
	private enum Days360Calctype
	{
		European,
		Us
	}
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		var numDate1 = ArgToDecimal(arguments, 0);
		var numDate2 = ArgToDecimal(arguments, 1);
		var dt1 = System.DateTime.FromOADate(numDate1);
		var dt2 = System.DateTime.FromOADate(numDate2);

		var calcType = Days360Calctype.Us;
		if (arguments.Count() > 2 && ArgToBool(arguments, 2))
		{
			calcType = Days360Calctype.European;
		}

		var (startDay, endDay) = calcType == Days360Calctype.European
			? AdjustDaysEuropean(dt1.Day, dt2.Day)
			: AdjustDaysUs(dt1, dt2);

		var result = (dt2.Year * 12 * 30 + dt2.Month * 30 + endDay) - (dt1.Year * 12 * 30 + dt1.Month * 30 + startDay);
		return CreateResult(result, DataType.Integer);
	}

	private static (int StartDay, int EndDay) AdjustDaysEuropean(int startDay, int endDay) =>
		(startDay == 31 ? 30 : startDay, endDay == 31 ? 30 : endDay);

	private static (int StartDay, int EndDay) AdjustDaysUs(System.DateTime dt1, System.DateTime dt2)
	{
		var startMonth = dt1.Month;
		var startDay = dt1.Day;
		var endMonth = dt2.Month;
		var endDay = dt2.Day;
		var nDaysInFeb = new GregorianCalendar().IsLeapYear(dt1.Year) ? 29 : 28;
		var startIsEndOfFebruary = startMonth == 2 && startDay == nDaysInFeb;

		// Both dates on the last day of February: D2 becomes 30.
		if (startIsEndOfFebruary && endMonth == 2 && endDay == nDaysInFeb)
		{
			endDay = 30;
		}

		if (startIsEndOfFebruary)
		{
			startDay = 30;
		}

		if (endDay == 31 && (startDay == 30 || startDay == 31))
		{
			endDay = 30;
		}

		if (startDay == 31)
		{
			startDay = 30;
		}

		return (startDay, endDay);
	}

	private static int GetNumWholeMonths(System.DateTime dt1, System.DateTime dt2)
	{
		var startDate = new System.DateTime(dt1.Year, dt1.Month, 1).AddMonths(1);
		var endDate = new System.DateTime(dt2.Year, dt2.Month, 1);
		return ((endDate.Year - startDate.Year) * 12) + (endDate.Month - startDate.Month);
	}
}

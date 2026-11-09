using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Yearfrac : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
		ValidateArguments(functionArguments, 2);
		var date1Num = ArgToDecimal(functionArguments, 0);
		var date2Num = ArgToDecimal(functionArguments, 1);
		if (date1Num > date2Num) //Switch to make date1 the lowest date
		{
			var t = date1Num;
			date1Num = date2Num;
			date2Num = t;
			var fa = functionArguments[1];
			functionArguments[1] = functionArguments[0];
			functionArguments[0] = fa;
		}

		var date1 = System.DateTime.FromOADate(date1Num);
		var date2 = System.DateTime.FromOADate(date2Num);

		var basis = GetBasis(functionArguments);

		var func = context.Configuration.FunctionRepository.GetFunction("days360");
		switch (basis)
		{
			case 0:
				return CreateResult(CalculateUsDays360Fraction(func, functionArguments, context, date1, date2), DataType.Decimal);
			case 1:
				return CreateResult(System.Math.Abs((date2 - date1).TotalDays / CalculateAcutalYear(date1, date2)), DataType.Decimal);
			case 2:
				return CreateResult(System.Math.Abs((date2 - date1).TotalDays / 360d), DataType.Decimal);
			case 3:
				return CreateResult(System.Math.Abs((date2 - date1).TotalDays / 365d), DataType.Decimal);
			case 4:
				var args = functionArguments.ToList();
				args.Add(new FunctionArgument(true));
				double? result = System.Math.Abs(func.Execute(args, context).ResultNumeric / 360d);
				return CreateResult(result.Value, DataType.Decimal);
			default:
				return null;
		}
	}

	private int GetBasis(FunctionArgument[] functionArguments)
	{
		var basis = 0;
		if (functionArguments.Length > 2)
		{
			basis = ArgToInt(functionArguments, 2);
			ThrowExcelErrorValueExceptionIf(() => basis is < 0 or > 4, eErrorType.Num);
		}

		return basis;
	}

	private static double CalculateUsDays360Fraction(ExcelFunction days360, FunctionArgument[] functionArguments, ParsingContext context, System.DateTime date1, System.DateTime date2)
	{
		var d360Result = System.Math.Abs(days360.Execute(functionArguments, context).ResultNumeric);
		// reproducing excels behaviour
		if (date1.Month == 2 && date2.Day == 31)
		{
			var daysInFeb = new GregorianCalendar().IsLeapYear(date1.Year) ? 29 : 28;
			if (date1.Day == daysInFeb)
			{
				d360Result++;
			}
		}

		return d360Result / 360d;
	}

	private static double CalculateAcutalYear(System.DateTime dt1, System.DateTime dt2)
	{
		var calendar = new GregorianCalendar();
		if (new System.DateTime(dt1.Year + 1, dt1.Month, dt1.Day) >= dt2)
		{
			return IncludesLeapDay(calendar, dt1, dt2) ? 366d : 365d;
		}

		var perYear = 0d;
		var nYears = dt2.Year - dt1.Year + 1;
		for (var y = dt1.Year; y <= dt2.Year; ++y)
		{
			perYear += calendar.IsLeapYear(y) ? 366 : 365;
		}

		return perYear / (double)nYears;
	}

	private static bool IncludesLeapDay(GregorianCalendar calendar, System.DateTime dt1, System.DateTime dt2) =>
		(calendar.IsLeapYear(dt1.Year) && dt1.Month <= 2)
		|| (calendar.IsLeapYear(dt2.Year) && dt2.Month > 2)
		|| (dt2.Month == 2 && dt2.Day == 29);
}

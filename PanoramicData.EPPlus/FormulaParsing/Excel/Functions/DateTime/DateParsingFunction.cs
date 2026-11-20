using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public abstract class DateParsingFunction : ExcelFunction
{
	private protected System.DateTime ParseFirstArgumentDate(IEnumerable<FunctionArgument> arguments)
	{
		ValidateArguments(arguments, 1);
		var dateObj = arguments.ElementAt(0).Value;
		return ParseDate(arguments, dateObj);
	}

	protected System.DateTime ParseDate(IEnumerable<FunctionArgument> arguments, object dateObj)
	{
		if (dateObj is string)
		{
			return System.DateTime.Parse(dateObj.ToString(), CultureInfo.InvariantCulture);
		}

		var d = ArgToDecimal(arguments, 0);
		return System.DateTime.FromOADate(d);
	}
}

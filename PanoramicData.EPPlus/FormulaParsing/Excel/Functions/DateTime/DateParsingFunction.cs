using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public abstract class DateParsingFunction : ExcelFunction
{
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

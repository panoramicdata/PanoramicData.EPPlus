using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

internal static class DatabaseRowCounter
{
	internal static int CountMatchingRows(IEnumerable<FunctionArgument> arguments, ParsingContext context, RowMatcher rowMatcher, Func<string> readField, Func<object, bool> shouldCountValue)
	{
		var dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
		string field = null;
		string criteriaRange;
		if (arguments.Count() == 2)
		{
			criteriaRange = arguments.ElementAt(1).ValueAsRangeInfo.Address.Address;
		}
		else
		{
			field = readField();
			criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;
		}

		var db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
		var criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);

		var nHits = 0;
		while (db.HasMoreRows)
		{
			var dataRow = db.Read();
			// Without a field name every matching row counts; with one, only rows whose field value qualifies.
			if (rowMatcher.IsMatch(dataRow, criteria) && (string.IsNullOrEmpty(field) || shouldCountValue(dataRow[field])))
			{
				nHits++;
			}
		}

		return nHits;
	}
}

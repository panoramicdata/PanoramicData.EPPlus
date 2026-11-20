/* Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2015-04-19
 *******************************************************************************/
using System.Diagnostics.CodeAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public abstract class DatabaseFunction : ExcelFunction
{
	protected RowMatcher RowMatcher { get; private set; }

	[SuppressMessage("Major Code Smell", "S3442:\"abstract\" classes should not have \"public\" constructors", Justification = "Published API: narrowing this constructor to protected is a breaking change for reflection and API compatibility.")]
	public DatabaseFunction()
		: this(new RowMatcher())
	{
	}

	[SuppressMessage("Major Code Smell", "S3442:\"abstract\" classes should not have \"public\" constructors", Justification = "Published API: narrowing this constructor to protected is a breaking change for reflection and API compatibility.")]
	public DatabaseFunction(RowMatcher rowMatcher)
	{
		RowMatcher = rowMatcher;
	}

	private protected CompileResult ExecuteAggregate(IEnumerable<FunctionArgument> arguments, ParsingContext context, System.Func<IEnumerable<double>, double> aggregate)
	{
		ValidateArguments(arguments, 3);
		var values = GetMatchingValues(arguments, context);
		return !values.Any() ? CreateResult(0d, DataType.Integer) : CreateResult(aggregate(values), DataType.Integer);
	}

	protected IEnumerable<double> GetMatchingValues(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		var dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
		var field = arguments.ElementAt(1).Value;
		var criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;

		var db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
		var criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);
		var values = new List<double>();

		while (db.HasMoreRows)
		{
			var dataRow = db.Read();
			if (!RowMatcher.IsMatch(dataRow, criteria))
			{
				continue;
			}

			var candidate = ConvertUtil.IsNumeric(field) ? dataRow[(int)ConvertUtil.GetValueDouble(field)] : dataRow[field.ToString().ToLower(CultureInfo.InvariantCulture)];
			if (ConvertUtil.IsNumeric(candidate))
			{
				values.Add(ConvertUtil.GetValueDouble(candidate));
			}
		}

		return values;
	}
}

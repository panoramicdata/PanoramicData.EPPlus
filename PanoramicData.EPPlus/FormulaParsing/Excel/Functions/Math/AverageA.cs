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
 * Mats Alm   		                Added		                2014-01-06
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class AverageA : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1, eErrorType.Div0);
		double nValues = 0d, result = 0d;
		foreach (var arg in arguments)
		{
			Calculate(arg, context, ref result, ref nValues);
		}

		return CreateResult(Divide(result, nValues), DataType.Decimal);
	}

	private void Calculate(FunctionArgument arg, ParsingContext context, ref double retVal, ref double nValues, bool isInArray = false)
	{
		if (ShouldIgnore(arg))
		{
			return;
		}

		if (arg.Value is IEnumerable<FunctionArgument> items)
		{
			foreach (var item in items)
			{
				Calculate(item, context, ref retVal, ref nValues, true);
			}
		}
		else if (arg.IsExcelRange)
		{
			AddRangeValues(arg.ValueAsRangeInfo, context, ref retVal, ref nValues);
		}
		else
		{
			AddSingleValue(arg.Value, isInArray, ref retVal, ref nValues);
		}

		CheckForAndHandleExcelError(arg);
	}

	private void AddRangeValues(ExcelDataProvider.IRangeInfo range, ParsingContext context, ref double retVal, ref double nValues)
	{
		foreach (var c in range)
		{
			if (ShouldIgnore(c, context))
			{
				continue;
			}

			CheckForAndHandleExcelError(c);
			if (IsNumeric(c.Value) && c.Value is not bool)
			{
				nValues++;
				retVal += c.ValueDouble;
			}
			else if (c.Value is bool boolValue)
			{
				nValues++;
				retVal += boolValue ? 1 : 0;
			}
			else if (c.Value is string)
			{
				nValues++;
			}
		}
	}

	private static void AddSingleValue(object value, bool isInArray, ref double retVal, ref double nValues)
	{
		var numericValue = GetAverageNumericValue(value, isInArray);
		if (numericValue.HasValue)
		{
			nValues++;
			retVal += numericValue.Value;
		}
		else if (value is string)
		{
			if (isInArray)
			{
				nValues++;
			}
			else
			{
				ThrowExcelErrorValueException(eErrorType.Value);
			}
		}
	}
}

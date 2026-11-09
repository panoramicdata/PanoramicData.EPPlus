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
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Sum : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		=> SumOverArguments(arguments, arg => Calculate(arg, context));

	private double Calculate(FunctionArgument arg, ParsingContext context)
	{
		var retVal = 0d;
		if (ShouldIgnore(arg))
		{
			return retVal;
		}

		if (arg.Value is IEnumerable<FunctionArgument> items)
		{
			foreach (var item in items)
			{
				retVal += Calculate(item, context);
			}
		}
		else if (arg.Value is ExcelDataProvider.IRangeInfo range)
		{
			retVal += SumVisibleCells(range, context, value => value);
		}
		else
		{
			CheckForAndHandleExcelError(arg);
			retVal += ConvertUtil.GetValueDouble(arg.Value, true);
		}

		return retVal;
	}
}

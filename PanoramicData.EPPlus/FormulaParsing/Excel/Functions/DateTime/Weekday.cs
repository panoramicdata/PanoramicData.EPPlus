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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Weekday : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		var serialNumber = ArgToDecimal(arguments, 0);
		var returnType = arguments.Count() > 1 ? ArgToInt(arguments, 1) : 1;
		return CreateResult(CalculateDayOfWeek(System.DateTime.FromOADate(serialNumber), returnType), DataType.Integer);
	}

	private static List<int> _oneBasedStartOnSunday = [1, 2, 3, 4, 5, 6, 7];
	private static List<int> _oneBasedStartOnMonday = [7, 1, 2, 3, 4, 5, 6];
	private static List<int> _zeroBasedStartOnSunday = [6, 0, 1, 2, 3, 4, 5];

	private static int CalculateDayOfWeek(System.DateTime dateTime, int returnType)
	{
		var dayIx = (int)dateTime.DayOfWeek;
		return returnType switch
		{
			1 => _oneBasedStartOnSunday[dayIx],
			2 => _oneBasedStartOnMonday[dayIx],
			3 => _zeroBasedStartOnSunday[dayIx],
			_ => throw new ExcelErrorValueException(eErrorType.Num),
		};
	}
}

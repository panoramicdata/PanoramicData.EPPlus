﻿/* Copyright (C) 2011  Jan Källman
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
 * Mats Alm   		                Added		                2015-01-15
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class ErrorType : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		var error = arguments.ElementAt(0);
		var isErrorFunc = context.Configuration.FunctionRepository.GetFunction("iserror");
		var isErrorResult = isErrorFunc.Execute(arguments, context);
		if (!(bool)isErrorResult.Result)
		{
			return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
		}

		var errorType = error.ValueAsExcelErrorValue;
		return errorType.Type switch
		{
			eErrorType.Null => CreateResult(1, DataType.Integer),
			eErrorType.Div0 => CreateResult(2, DataType.Integer),
			eErrorType.Value => CreateResult(3, DataType.Integer),
			eErrorType.Ref => CreateResult(4, DataType.Integer),
			eErrorType.Name => CreateResult(5, DataType.Integer),
			eErrorType.Num => CreateResult(6, DataType.Integer),
			eErrorType.NA => CreateResult(7, DataType.Integer),
			_ => CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError),
		};
	}
}

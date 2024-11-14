/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExpressionConverter : IExpressionConverter
{
	public StringExpression ToStringExpression(Expression expression)
	{
		var result = expression.Compile();
		var newExp = new StringExpression(result.Result.ToString())
		{
			Operator = expression.Operator
		};
		return newExp;
	}

	public Expression FromCompileResult(CompileResult compileResult) => compileResult.DataType switch
	{
		DataType.Integer => compileResult.Result is string
							? new IntegerExpression(compileResult.Result.ToString())
							: new IntegerExpression(Convert.ToDouble(compileResult.Result)),
		DataType.String => new StringExpression(compileResult.Result.ToString()),
		DataType.Decimal => compileResult.Result is string
								   ? new DecimalExpression(compileResult.Result.ToString())
								   : new DecimalExpression(((double)compileResult.Result)),
		DataType.Boolean => compileResult.Result is string
								   ? new BooleanExpression(compileResult.Result.ToString())
								   : new BooleanExpression((bool)compileResult.Result),
		//case DataType.Enumerable:
		//    return 
		DataType.ExcelError => compileResult.Result is string
							? new ExcelErrorExpression(compileResult.Result.ToString(),
								ExcelErrorValue.Parse(compileResult.Result.ToString()))
							: new ExcelErrorExpression((ExcelErrorValue)compileResult.Result),//throw (new OfficeOpenXml.FormulaParsing.Exceptions.ExcelErrorValueException((ExcelErrorValue)compileResult.Result)); //Added JK
		DataType.Empty => new IntegerExpression(0),//Added JK
		DataType.Time or DataType.Date => new DecimalExpression((double)compileResult.Result),
		_ => null,
	};

	private static IExpressionConverter _instance;
	public static IExpressionConverter Instance
	{
		get
		{
			_instance ??= new ExpressionConverter();

			return _instance;
		}
	}
}

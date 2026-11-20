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

public class CompileResultFactory
{
	public virtual CompileResult Create(object obj)
	{
		if (obj is ExcelDataProvider.INameInfo nameInfo)
		{
			obj = nameInfo.Value;
		}

		if (obj is ExcelDataProvider.IRangeInfo rangeInfo)
		{
			obj = rangeInfo.GetOffset(0, 0);
		}

		if (obj == null)
		{
			return new CompileResult(null, DataType.Empty);
		}

		if (obj is System.DateTime dateTime)
		{
			return new CompileResult(dateTime.ToOADate(), DataType.Date);
		}

		var dataType = GetDataType(obj);
		return dataType.HasValue
			? new CompileResult(obj, dataType.Value)
			: throw new ArgumentException("Non supported type " + obj.GetType().FullName);
	}

	private static DataType? GetDataType(object obj) => obj switch
	{
		string => DataType.String,
		double or decimal => DataType.Decimal,
		int or long or short => DataType.Integer,
		bool => DataType.Boolean,
		_ when obj.GetType() == typeof(ExcelErrorValue) => DataType.ExcelError,
		_ => null
	};
}

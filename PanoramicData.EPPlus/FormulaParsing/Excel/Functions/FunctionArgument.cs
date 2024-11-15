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
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class FunctionArgument
{
	public FunctionArgument(object? val)
	{
		Value = val;
		DataType = DataType.Unknown;
	}

	public FunctionArgument(object val, DataType dataType)
		: this(val)
	{
		DataType = dataType;
	}

	private ExcelCellState _excelCellState;

	public void SetExcelStateFlag(ExcelCellState state) => _excelCellState |= state;

	public bool ExcelStateFlagIsSet(ExcelCellState state) => (_excelCellState & state) != 0;

	public object? Value { get; private set; }

	public DataType DataType { get; }

	public Type Type => Value?.GetType();

	public int ExcelAddressReferenceId { get; set; }

	public bool IsExcelRange => Value is not null and ExcelDataProvider.IRangeInfo;

	public bool ValueIsExcelError => ExcelErrorValue.Values.IsErrorValue(Value);

	public ExcelErrorValue ValueAsExcelErrorValue => ExcelErrorValue.Parse(Value.ToString());

	public EpplusExcelDataProvider.IRangeInfo ValueAsRangeInfo => Value as EpplusExcelDataProvider.IRangeInfo;
	public object ValueFirst
	{
		get
		{
			if (Value is ExcelDataProvider.INameInfo info)
			{
				Value = info.Value;
			}

			return Value is not ExcelDataProvider.IRangeInfo v ? Value : v.GetValue(v.Address._fromRow, v.Address._fromCol);
		}
	}

}

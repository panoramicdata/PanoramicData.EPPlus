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
using System.Linq;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class CompileResult
{
	private static CompileResult _empty = new(null, DataType.Empty);

	public static CompileResult Empty => _empty;

	private double? _ResultNumeric;

	public CompileResult(object result, DataType dataType)
	{
		if (result is ExcelDoubleCellValue)
		{
			Result = ((ExcelDoubleCellValue)result).Value;
		}
		else
		{
			Result = result;
		}

		DataType = dataType;
	}

	public CompileResult(eErrorType errorType)
	{
		Result = ExcelErrorValue.Create(errorType);
		DataType = DataType.ExcelError;
	}

	public CompileResult(ExcelErrorValue errorValue)
	{
		Require.Argument(errorValue).IsNotNull("errorValue");
		Result = errorValue;
		DataType = DataType.ExcelError;
	}

	public object Result
	{
		get;
		private set;
	}

	public object ResultValue
	{
		get
		{
			var r = Result as ExcelDataProvider.IRangeInfo;
			return r == null ? Result : r.GetValue(r.Address._fromRow, r.Address._fromCol);
		}
	}

	public double ResultNumeric
	{
		get
		{
			// We assume that Result does not change unless it is a range.
			if (_ResultNumeric == null)
			{
				if (IsNumeric)
				{
					_ResultNumeric = Result == null ? 0 : Convert.ToDouble(Result);
				}
				else if (Result is DateTime)
				{
					_ResultNumeric = ((DateTime)Result).ToOADate();
				}
				else if (Result is TimeSpan)
				{
					_ResultNumeric = DateTime.FromOADate(0).Add((TimeSpan)Result).ToOADate();
				}
				else if (Result is ExcelDataProvider.IRangeInfo)
				{
					var c = ((ExcelDataProvider.IRangeInfo)Result).FirstOrDefault();
					return c == null ? 0 : c.ValueDoubleLogical;
				}
				// The IsNumericString and IsDateString properties will set _ResultNumeric for efficiency so we just need
				// to check them here.
				else if (!IsNumericString && !IsDateString)
				{
					_ResultNumeric = 0;
				}
			}

			return _ResultNumeric.Value;
		}
	}

	public DataType DataType
	{
		get;
		private set;
	}

	public bool IsNumeric => DataType is DataType.Decimal or DataType.Integer or DataType.Empty or DataType.Boolean or DataType.Date;

	public bool IsNumericString
	{
		get
		{
			if (DataType == DataType.String && ConvertUtil.TryParseNumericString(Result, out var result))
			{
				_ResultNumeric = result;
				return true;
			}

			return false;
		}
	}

	public bool IsDateString
	{
		get
		{
			if (DataType == DataType.String && ConvertUtil.TryParseDateString(Result, out var result))
			{
				_ResultNumeric = result.ToOADate();
				return true;
			}

			return false;
		}
	}

	public bool IsResultOfSubtotal { get; set; }

	public bool IsHiddenCell { get; set; }

	public int ExcelAddressReferenceId { get; set; }

	public bool IsResultOfResolvedExcelRange => ExcelAddressReferenceId > 0;
}

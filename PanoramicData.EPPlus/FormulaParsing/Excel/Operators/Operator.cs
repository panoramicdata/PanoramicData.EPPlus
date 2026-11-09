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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators;

[System.Diagnostics.CodeAnalysis.SuppressMessage("Major Code Smell", "S3453:Classes should not have only private constructors", Justification = "Instances are created internally and exposed through the static operator properties.")]
public class Operator : IOperator
{
	private const int PrecedencePercent = 2;
	private const int PrecedenceExp = 4;
	private const int PrecedenceMultiplyDevide = 6;
	private const int PrecedenceIntegerDivision = 8;
	private const int PrecedenceModulus = 10;
	private const int PrecedenceAddSubtract = 12;
	private const int PrecedenceConcat = 15;
	private const int PrecedenceComparison = 25;

	private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, CompileResult> implementation)
	{
		_implementation = implementation;
		_precedence = precedence;
		_operator = @operator;
	}

	private readonly Func<CompileResult, CompileResult, CompileResult> _implementation;
	private readonly int _precedence;
	private readonly Operators _operator;

	int IOperator.Precedence => _precedence;

	Operators IOperator.Operator => _operator;

	public CompileResult Apply(CompileResult left, CompileResult right)
	{
		if (left.Result is ExcelErrorValue)
		{
			return new CompileResult(left.Result, DataType.ExcelError);
		}
		else if (right.Result is ExcelErrorValue)
		{
			return new CompileResult(right.Result, DataType.ExcelError);
		}

		return _implementation(left, right);
	}

	public override string ToString() => "Operator: " + _operator;

	private static IOperator _plus;
	public static IOperator Plus => _plus ??= new Operator(Operators.Plus, PrecedenceAddSubtract, (l, r) =>
	{
		l = ZeroIfNullResult(l);
		r = ZeroIfNullResult(r);
		return EitherIsError(l, r, out var errorVal)
			? new CompileResult(errorVal)
			: ApplyArithmetic(l, r, (x, y) => x + y);
	});

	private static IOperator _minus;
	public static IOperator Minus => _minus ??= new Operator(Operators.Minus, PrecedenceAddSubtract,
		(l, r) => ApplyArithmetic(ZeroIfNullResult(l), ZeroIfNullResult(r), (x, y) => x - y));

	private static IOperator _multiply;
	public static IOperator Multiply => _multiply ??= new Operator(Operators.Multiply, PrecedenceMultiplyDevide,
		(l, r) => ApplyArithmetic(l ?? Zero(), r ?? Zero(), (x, y) => x * y));

	private static IOperator _divide;
	public static IOperator Divide => _divide ??= new Operator(Operators.Divide, PrecedenceMultiplyDevide, (l, r) =>
	{
		if (!IsNumericLike(l) || !IsNumericLike(r))
		{
			return new CompileResult(eErrorType.Value);
		}

		var left = l.ResultNumeric;
		var right = r.ResultNumeric;
		return Math.Abs(right - 0d) < double.Epsilon
			? new CompileResult(eErrorType.Div0)
			: new CompileResult(left / right, DataType.Decimal);
	});

	public static IOperator Exp => new Operator(Operators.Exponentiation, PrecedenceExp, (l, r) =>
	{
		if (l == null && r == null)
		{
			return new CompileResult(eErrorType.Value);
		}

		l ??= Zero();
		r ??= Zero();
		return IsNumericLike(l) && IsNumericLike(r)
			? new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), DataType.Decimal)
			: new CompileResult(0d, DataType.Decimal);
	});

	public static IOperator Concat => new Operator(Operators.Concat, PrecedenceConcat, (l, r) =>
	{
		l ??= new CompileResult(string.Empty, DataType.String);
		r ??= new CompileResult(string.Empty, DataType.String);
		var lStr = l.Result != null ? l.ResultValue.ToString() : string.Empty;
		var rStr = r.Result != null ? r.ResultValue.ToString() : string.Empty;
		return new CompileResult(string.Concat(lStr, rStr), DataType.String);
	});

	private static IOperator _greaterThan;
	public static IOperator GreaterThan => _greaterThan ??= CreateComparison(Operators.GreaterThan, compRes => compRes > 0);

	private static IOperator _eq;
	public static IOperator Eq => _eq ??= CreateComparison(Operators.Equals, compRes => compRes == 0);

	private static IOperator _notEqualsTo;
	public static IOperator NotEqualsTo => _notEqualsTo ??= CreateComparison(Operators.NotEqualTo, compRes => compRes != 0);

	private static IOperator _greaterThanOrEqual;
	public static IOperator GreaterThanOrEqual => _greaterThanOrEqual ??= CreateComparison(Operators.GreaterThanOrEqual, compRes => compRes >= 0);

	private static IOperator _lessThan;
	public static IOperator LessThan => _lessThan ??= CreateComparison(Operators.LessThan, compRes => compRes < 0);

	public static IOperator LessThanOrEqual => CreateComparison(Operators.LessThanOrEqual, compRes => compRes <= 0);

	private static IOperator _percent;
	public static IOperator Percent => _percent ??= new Operator(Operators.Percent, PrecedencePercent,
		(l, r) => ApplyArithmetic(l ?? Zero(), r ?? Zero(), (x, y) => x * y));

	private static Operator CreateComparison(Operators @operator, Func<int, bool> comparison) =>
		new(@operator, PrecedenceComparison, (l, r) => Compare(l, r, comparison));

	private static CompileResult Zero() => new(0, DataType.Integer);

	private static CompileResult ZeroIfNullResult(CompileResult result) => result == null || result.Result == null ? Zero() : result;

	// Evaluated in this order because IsNumericString and IsDateString cache ResultNumeric as a side effect.
	private static bool IsNumericLike(CompileResult result) =>
		result.IsNumeric || result.IsNumericString || result.IsDateString || result.Result is ExcelDataProvider.IRangeInfo;

	private static CompileResult ApplyArithmetic(CompileResult l, CompileResult r, Func<double, double, double> operation)
	{
		if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
		{
			return new CompileResult(operation(l.ResultNumeric, r.ResultNumeric), DataType.Integer);
		}

		if (IsNumericLike(l) && IsNumericLike(r))
		{
			return new CompileResult(operation(l.ResultNumeric, r.ResultNumeric), DataType.Decimal);
		}

		return new CompileResult(eErrorType.Value);
	}

	private static object GetObjFromOther(CompileResult obj, CompileResult other)
	{
		if (obj.Result == null)
		{
			return other.DataType == DataType.String ? string.Empty : 0d;
		}

		return obj.ResultValue;
	}

	private static CompileResult Compare(CompileResult l, CompileResult r, Func<int, bool> comparison)
	{
		if (EitherIsError(l, r, out var errorVal))
		{
			return new CompileResult(errorVal);
		}

		object left, right;
		left = GetObjFromOther(l, r);
		right = GetObjFromOther(r, l);
		if (ConvertUtil.IsNumeric(left) && ConvertUtil.IsNumeric(right))
		{
			var lnum = ConvertUtil.GetValueDouble(left);
			var rnum = ConvertUtil.GetValueDouble(right);
			if (Math.Abs(lnum - rnum) < double.Epsilon)
			{
				return new CompileResult(comparison(0), DataType.Boolean);
			}

			var comparisonResult = lnum.CompareTo(rnum);
			return new CompileResult(comparison(comparisonResult), DataType.Boolean);
		}
		else
		{
			var comparisonResult = CompareString(left, right);
			return new CompileResult(comparison(comparisonResult), DataType.Boolean);
		}
	}

	private static int CompareString(object l, object r)
	{
		var sl = (l ?? "").ToString();
		var sr = (r ?? "").ToString();
		return System.String.Compare(sl, sr, StringComparison.OrdinalIgnoreCase);
	}

	private static bool EitherIsError(CompileResult l, CompileResult r, out ExcelErrorValue errorVal)
	{
		if (l.DataType == DataType.ExcelError)
		{
			errorVal = (ExcelErrorValue)l.Result;
			return true;
		}

		if (r.DataType == DataType.ExcelError)
		{
			errorVal = (ExcelErrorValue)r.Result;
			return true;
		}

		errorVal = null;
		return false;
	}
}

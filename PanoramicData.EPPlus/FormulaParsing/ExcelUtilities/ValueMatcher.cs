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

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class ValueMatcher
{
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S2339:Public constant members should not be used", Justification = "Public API of a published library; changing to static readonly would break consumers compiled against the constant.")]
	public const int IncompatibleOperands = -2;

	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S1449:Culture should be specified for string operations", Justification = "Matching uses the current culture by design; specifying another culture would change results.")]
	public virtual int IsMatch(object o1, object o2)
	{
		var nullComparison = CompareNulls(o1, o2);
		if (nullComparison.HasValue)
		{
			return nullComparison.Value;
		}
		//Handle ranges and defined names
		o1 = CheckGetRange(o1);
		o2 = CheckGetRange(o2);

		if (o1 is string && o2 is string)
		{
			return CompareStringToString(o1.ToString().ToLower(), o2.ToString().ToLower());
		}
		else if (o1.GetType() == typeof(string))
		{
			return CompareStringToObject(o1.ToString(), o2);
		}
		else if (o2.GetType() == typeof(string))
		{
			return CompareObjectToString(o1, o2.ToString());
		}

		return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
	}

	private static int? CompareNulls(object o1, object o2)
	{
		if (o1 == null)
		{
			return o2 == null ? 0 : -1;
		}

		return o2 == null ? 1 : null;
	}

	private static object CheckGetRange(object v)
	{
		if (v is ExcelDataProvider.IRangeInfo r)
		{
			v = r.GetOffset(0, 0);
		}
		else if (v is ExcelDataProvider.INameInfo n)
		{
			v = CheckGetRange(n);
		}

		return v;
	}

	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S1449:Culture should be specified for string operations", Justification = "Comparison uses the current culture by design; specifying another culture would change results.")]
	protected virtual int CompareStringToString(string s1, string s2) => s1.CompareTo(s2);

	protected virtual int CompareStringToObject(string o1, object o2)
	{
		if (double.TryParse(o1, out var d1))
		{
			return d1.CompareTo(Convert.ToDouble(o2));
		}

		if (bool.TryParse(o1, out var b1))
		{
			return b1.CompareTo(Convert.ToBoolean(o2));
		}

		return DateTime.TryParse(o1, out var dt1) ? dt1.CompareTo(Convert.ToDateTime(o2)) : IncompatibleOperands;
	}

	protected virtual int CompareObjectToString(object o1, string o2) => double.TryParse(o2, out var d2) ? Convert.ToDouble(o1).CompareTo(d2) : IncompatibleOperands;
}

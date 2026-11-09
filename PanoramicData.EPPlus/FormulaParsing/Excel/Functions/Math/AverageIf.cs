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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Linq;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class AverageIf : HiddenValuesHandlingFunction
{
	private readonly ExpressionEvaluator _expressionEvaluator;

	public AverageIf()
		: this(new ExpressionEvaluator())
	{

	}

	public AverageIf(ExpressionEvaluator evaluator)
	{
		Require.That(evaluator).Named("evaluator").IsNotNull();
		_expressionEvaluator = evaluator;
	}

	private bool Evaluate(object obj, string expression)
	{
		var candidate = default(double?);
		if (IsNumeric(obj))
		{
			candidate = ConvertUtil.GetValueDouble(obj);
		}

		return candidate.HasValue
			? _expressionEvaluator.Evaluate(candidate.Value, expression)
			: _expressionEvaluator.Evaluate(obj, expression);
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		var criteria = arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;
		var retVal = 0d;
		if (arguments.ElementAt(0).Value is not ExcelDataProvider.IRangeInfo args)
		{
			var val = arguments.ElementAt(0).Value;
			if (criteria != null && Evaluate(val, criteria))
			{
				var lookupRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				retVal = arguments.Count() > 2
					? lookupRange.First().ValueDouble
					: ConvertUtil.GetValueDouble(val, true);
			}
			else
			{
				throw new ExcelErrorValueException(eErrorType.Div0);
			}
		}
		else if (arguments.Count() > 2)
		{
			var lookupRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
			retVal = CalculateWithLookupRange(args, criteria, lookupRange);
		}
		else
		{
			retVal = CalculateSingleRange(args, criteria);
		}

		return CreateResult(retVal, DataType.Decimal);
	}

	private double CalculateWithLookupRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange)
	{
		var retVal = CriteriaRangeSum.SumMatchingOffsets(range, sumRange, value => criteria != null && Evaluate(value, criteria), out var nMatches);
		return Divide(retVal, nMatches);
	}

	private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression)
	{
		var retVal = CriteriaRangeSum.SumMatchingCells(range, value => expression != null && IsNumeric(value) && Evaluate(value, expression), out var nMatches);
		return Divide(retVal, nMatches);
	}
}

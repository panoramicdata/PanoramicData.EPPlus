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
 * Mats Alm   		                Added		                2015-01-15
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class SumIfs : MultipleRangeCriteriasFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
		ValidateArguments(functionArguments, 3);
		var valueRange = functionArguments[0].ValueAsRangeInfo;
		List<double> sumRange;
		if (valueRange != null)
		{
			sumRange = ArgsToDoubleEnumerableZeroPadded(true, valueRange, context).ToList();
		}
		else
		{
			sumRange = ArgsToDoubleEnumerable(true, new List<FunctionArgument> { functionArguments[0] }, context).Select(x => (double)x).ToList();
		}

		var matchIndexes = GetMatchIndexesForCriteria(functionArguments, 1);

		var result = matchIndexes.Sum(index => sumRange[index]);

		return CreateResult(result, DataType.Decimal);
	}
}

using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Rank : ExcelFunction
{
	readonly bool _isAvg;
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Major Code Smell", "S2360:Optional parameters should not be used", Justification = "Public API of a published library; changing it would break existing consumers.")]
	public Rank(bool isAvg = false)
	{
		_isAvg = isAvg;
	}
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		var number = ArgToDecimal(arguments, 0);
		var refer = arguments.ElementAt(1);
		var asc = false;
		if (arguments.Count() > 2)
		{
			asc = ArgToBool(arguments, 2);
		}

		var l = GetSortedNumbers(refer.ValueAsRangeInfo);
		var ix = asc ? GetAscendingRank(l, number) : GetDescendingRank(l, number);

		return ix <= 0 || ix > l.Count
			? new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError)
			: CreateResult(ix, DataType.Decimal);
	}

	private static List<double> GetSortedNumbers(ExcelDataProvider.IRangeInfo range)
	{
		var l = new List<double>();
		foreach (var c in range)
		{
			var v = Utils.ConvertUtil.GetValueDouble(c.Value, false, true);
			if (!double.IsNaN(v))
			{
				l.Add(v);
			}
		}

		l.Sort();
		return l;
	}

	private double GetAscendingRank(List<double> l, double number)
	{
		double ix = l.IndexOf(number) + 1;
		if (_isAvg)
		{
			var st = Convert.ToInt32(ix);
			while (l.Count > st && l[st] == number)
			{
				st++;
			}

			if (st > ix)
			{
				ix += (st - ix) / 2D;
			}
		}

		return ix;
	}

	private double GetDescendingRank(List<double> l, double number)
	{
		double ix = l.LastIndexOf(number);
		if (_isAvg)
		{
			var st = Convert.ToInt32(ix) - 1;
			while (0 <= st && l[st] == number)
			{
				st--;
			}

			if (st + 1 < ix)
			{
				ix -= (ix - st - 1) / 2D;
			}
		}

		return l.Count - ix;
	}
}

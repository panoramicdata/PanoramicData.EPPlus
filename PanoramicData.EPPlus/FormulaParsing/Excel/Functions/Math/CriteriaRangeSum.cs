using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

internal static class CriteriaRangeSum
{
	internal static double SumMatchingOffsets(ExcelDataProvider.IRangeInfo range, ExcelDataProvider.IRangeInfo sumRange, Func<object, bool> isMatch, out int nMatches)
	{
		var retVal = 0d;
		nMatches = 0;
		foreach (var cell in range)
		{
			if (isMatch(cell.Value))
			{
				var or = cell.Row - range.Address._fromRow;
				var oc = cell.Column - range.Address._fromCol;
				if (sumRange.Address._fromRow + or <= sumRange.Address._toRow &&
				   sumRange.Address._fromCol + oc <= sumRange.Address._toCol)
				{
					var v = sumRange.GetOffset(or, oc);
					if (v is ExcelErrorValue errorValue)
					{
						throw new ExcelErrorValueException(errorValue);
					}

					nMatches++;
					retVal += ConvertUtil.GetValueDouble(v, true);
				}
			}
		}

		return retVal;
	}

	internal static double SumMatchingCells(ExcelDataProvider.IRangeInfo range, Func<object, bool> isMatch, out int nMatches)
	{
		var retVal = 0d;
		nMatches = 0;
		foreach (var candidate in range)
		{
			if (isMatch(candidate.Value))
			{
				if (candidate.IsExcelError)
				{
					throw new ExcelErrorValueException((ExcelErrorValue)candidate.Value);
				}

				retVal += candidate.ValueDouble;
				nMatches++;
			}
		}

		return retVal;
	}
}

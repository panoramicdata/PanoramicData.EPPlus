using OfficeOpenXml.Table;
using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	internal void SetTableTotalFunction(ExcelTable tbl, ExcelTableColumn col, int colNum = -1)
	{
		if (tbl.ShowTotal == false)
		{
			return;
		}

		if (colNum == -1)
		{
			for (var i = 0; i < tbl.Columns.Count; i++)
			{
				if (tbl.Columns[i].Name == col.Name)
				{
					colNum = tbl.Address._fromCol + i;
				}
			}
		}

		if (col.TotalsRowFunction == RowFunctions.Custom)
		{
			SetFormula(tbl.Address._toRow, colNum, col.TotalsRowFormula);
		}
		else if (col.TotalsRowFunction != RowFunctions.None)
		{
			var functionNumber = GetSubtotalFunctionNumber(col.TotalsRowFunction);
			SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, functionNumber));

		}
		else
		{
			SetValueInner(tbl.Address._toRow, colNum, col.TotalsRowLabel);

		}
	}

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	private static string GetSubtotalFunctionNumber(RowFunctions totalsRowFunction) => totalsRowFunction switch
	{
		RowFunctions.Average => "101",
		RowFunctions.CountNums => "102",
		RowFunctions.Count => "103",
		RowFunctions.Max => "104",
		RowFunctions.Min => "105",
		RowFunctions.StdDev => "107",
		RowFunctions.Var => "110",
		RowFunctions.Sum => "109",
		_ => throw new Exception("Unknown RowFunction enum")
	};

	internal void SetFormula(int row, int col, object value)
	{
		_formulas.SetValue(row, col, value);
		if (!ExistsValueInner(row, col))
		{
			SetValueInner(row, col, null);
		}
	}
}

using OfficeOpenXml.Table;
using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
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
			switch (col.TotalsRowFunction)
			{
				case RowFunctions.Average:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "101"));
					break;
				case RowFunctions.CountNums:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "102"));
					break;
				case RowFunctions.Count:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "103"));
					break;
				case RowFunctions.Max:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "104"));
					break;
				case RowFunctions.Min:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "105"));
					break;
				case RowFunctions.StdDev:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "107"));
					break;
				case RowFunctions.Var:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "110"));
					break;
				case RowFunctions.Sum:
					SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "109"));
					break;
				default:
					throw new Exception("Unknown RowFunction enum");
			}
		}
		else
		{
			SetValueInner(tbl.Address._toRow, colNum, col.TotalsRowLabel);

		}
	}

	internal void SetFormula(int row, int col, object value)
	{
		_formulas.SetValue(row, col, value);
		if (!ExistsValueInner(row, col))
		{
			SetValueInner(row, col, null);
		}
	}
}

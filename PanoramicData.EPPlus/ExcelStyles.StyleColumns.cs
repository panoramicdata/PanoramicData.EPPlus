using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	private void SetStyleColumns(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
	{
		SetStyleColumnRange(sender, e, address, ws, styleCashe);

		//Set for individual cells in the span. We loop all cells here since the cells are sorted with columns first.
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, address._fromCol, address._toRow, address._toCol);
		while (cse.Next())
		{
			if (cse.Column >= address.Start.Column &&
				cse.Column <= address.End.Column &&
				cse.Value._styleId != 0)
			{
				ws.SetStyleInner(cse.Row, cse.Column, GetOrAddStyleId(sender, e, styleCashe, cse.Value._styleId));
			}
		}

		if (!(address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns))
		{
			SetStyleCellsInStyledColumns(sender, e, address, ws, styleCashe);
		}
	}

	private void SetStyleColumnRange(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
	{
		ExcelColumn column;
		int col = address.Start.Column, row = 0;
		bool isNew;
		//Get the startcolumn
		object o = null;
		if (!ws.ExistsValueInner(0, address.Start.Column, ref o))
		{
			column = ws.Column(address.Start.Column);
			isNew = true;
		}
		else
		{
			column = (ExcelColumn)o;
			isNew = false;
		}

		var prevColumMax = column.ColumnMax;
		while (column.ColumnMin <= address.End.Column)
		{
			if (column.ColumnMin > prevColumMax + 1)
			{
				var newColumn = ws.Column(prevColumMax + 1);
				newColumn.ColumnMax = column.ColumnMin - 1;
				AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
			}

			if (column.ColumnMax > address.End.Column)
			{
				ws.CopyColumn(column, address.End.Column + 1, column.ColumnMax);
				column.ColumnMax = address.End.Column;
			}

			var s = ws.GetStyleInner(0, column.ColumnMin);
			AddNewStyleColumn(sender, e, ws, styleCashe, column, s);

			prevColumMax = column.ColumnMax;
			if (!ws._values.NextCell(ref row, ref col) || row > 0)
			{
				column = CloseColumnRange(sender, e, address, ws, styleCashe, column, isNew);
				break;
			}
			else
			{
				column = (ws.GetValueInner(0, col) as ExcelColumn);
			}
		}

		if (column._columnMax < address.End.Column)
		{
			var newCol = ws.Column(column._columnMax + 1) as ExcelColumn;
			newCol._columnMax = address.End.Column;

			var s = ws.GetStyleInner(0, column.ColumnMin);
			ws.SetStyleInner(0, column.ColumnMin, GetOrAddStyleId(sender, e, styleCashe, s));

			column._columnMax = address.End.Column;
		}
	}

	private ExcelColumn CloseColumnRange(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, bool isNew)
	{
		if (column._columnMax == address.End.Column)
		{
			return column;
		}

		if (isNew)
		{
			column._columnMax = address.End.Column;
		}
		else
		{
			var newColumn = ws.Column(column._columnMax + 1);
			newColumn.ColumnMax = address.End.Column;
			AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
			column = newColumn;
		}

		return column;
	}

	private void SetStyleCellsInStyledColumns(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
	{
		//Update cells with styled columns
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, 0, address._toRow, 0);
		while (cse.Next())
		{
			if (cse.Value._styleId == 0)
			{
				continue;
			}

			for (var c = address._fromCol; c <= address._toCol; c++)
			{
				if (!ws.ExistsStyleInner(cse.Row, c))
				{
					ws.SetStyleInner(cse.Row, c, GetOrAddStyleId(sender, e, styleCashe, cse.Value._styleId));
				}
			}
		}
	}

	private void AddNewStyleColumn(StyleBase sender, StyleChangeEventArgs e, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, int s) => ws.SetStyleInner(0, column.ColumnMin, GetOrAddStyleId(sender, e, styleCashe, s));
}

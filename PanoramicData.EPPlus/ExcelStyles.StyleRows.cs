using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	private void SetStyleRows(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
	{
		for (var rowNum = address.Start.Row; rowNum <= address.End.Row; rowNum++)
		{
			var s = ws.GetStyleInner(rowNum, 0);
			if (s == 0)
			{
				s = SetRowStyleFromColumns(ws, rowNum);
			}

			ws.SetStyleInner(rowNum, 0, GetOrAddStyleId(sender, e, styleCashe, s));
		}

		//Update individual cells
		var cse2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);
		while (cse2.Next())
		{
			var s = cse2.Value._styleId;
			if (s == 0)
			{
				continue;
			}

			ws.SetStyleInner(cse2.Row, cse2.Column, GetOrAddStyleId(sender, e, styleCashe, s));
		}

		SetStyleCellsInStyledRows(sender, e, address, ws, styleCashe);
	}

	private static int SetRowStyleFromColumns(ExcelWorksheet ws, int rowNum)
	{
		var s = 0;
		//iterate all columns and set the row to the style of the last column
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
		while (cse.Next())
		{
			s = cse.Value._styleId;
			if (s == 0)
			{
				continue;
			}

			if (ws.GetValueInner(cse.Row, cse.Column) is ExcelColumn c && c.ColumnMax < ExcelPackage.MaxColumns)
			{
				for (var col = c.ColumnMin; col < c.ColumnMax; col++)
				{
					if (!ws.ExistsStyleInner(rowNum, col))
					{
						ws.SetStyleInner(rowNum, col, s);
					}
				}
			}
		}

		ws.SetStyleInner(rowNum, 0, s);
		cse.Dispose();
		return s;
	}

	private void SetStyleCellsInStyledRows(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
	{
		//Update cells with styled rows
		var cse2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, address._toCol);
		while (cse2.Next())
		{
			if (cse2.Value._styleId == 0)
			{
				continue;
			}

			for (var r = address._fromRow; r <= address._toRow; r++)
			{
				if (!ws.ExistsStyleInner(r, cse2.Column))
				{
					var s = cse2.Value._styleId;
					ws.SetStyleInner(r, cse2.Column, GetOrAddStyleId(sender, e, styleCashe, s));
				}
			}
		}
	}
}

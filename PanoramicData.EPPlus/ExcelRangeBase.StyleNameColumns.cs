namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private void SetStyleNameOnColumns(string value)
	{
		ExcelColumn column;
		var c = _worksheet.GetValue(0, _fromCol);
		if (c == null)
		{
			column = _worksheet.Column(_fromCol);
		}
		else
		{
			column = (ExcelColumn)c;
		}

		column.StyleName = value;
		column.StyleID = _styleID;

		var cols = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, _fromCol + 1, 0, _toCol);
		if (cols.Next())
		{
			column = SetStyleNameOnFollowingColumns(column, cols, value);
		}

		if (column.ColumnMax < _toCol)
		{
			column.ColumnMax = _toCol;
		}

		if (IsFullRow()) //FullRow
		{
			SetStyleIdOnRowHeaders();
		}
	}

	private ExcelColumn SetStyleNameOnFollowingColumns(ExcelColumn column, CellsStoreEnumerator<ExcelCoreValue> cols, string value)
	{
		while (column.ColumnMin <= _toCol)
		{
			if (column.ColumnMax > _toCol)
			{
				_worksheet.CopyColumn(column, _toCol + 1, column.ColumnMax);
				column.ColumnMax = _toCol;
			}

			column._styleName = value;
			column.StyleID = _styleID;

			if (cols.Value._value == null)
			{
				break;
			}
			else
			{
				var nextCol = (ExcelColumn)cols.Value._value;
				if (column.ColumnMax < nextCol.ColumnMax - 1)
				{
					column.ColumnMax = nextCol.ColumnMax - 1;
				}

				column = nextCol;
				cols.Next();
			}
		}

		return column;
	}

	private void SetStyleIdOnRowHeaders()
	{
		var rows = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 1, 0, ExcelPackage.MaxRows, 0);
		rows.Next();
		while (rows.Value._value != null)
		{
			_worksheet.SetStyleInner(rows.Row, 0, _styleID);
			if (!rows.Next())
			{
				break;
			}
		}
	}

	private void SetStyleIdOnEachCell()
	{
		for (var c = _fromCol; c <= _toCol; c++)
		{
			for (var r = _fromRow; r <= _toRow; r++)
			{
				_worksheet.SetStyleInner(r, c, _styleID);
			}
		}
	}
}

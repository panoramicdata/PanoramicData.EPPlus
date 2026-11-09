using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// The styleobject for the range.
	/// </summary>
	public ExcelStyle Style
	{
		get
		{
			IsRangeValid("styling");
			var s = 0;
			if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref s)) //Cell exists
			{
				if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref s)) //No, check Row style
				{
					var c = Worksheet.GetColumn(_fromCol);
					if (c == null)
					{
						s = 0;
					}
					else
					{
						s = c.StyleID;
					}
				}
			}

			return _worksheet.Workbook.Styles.GetStyleObject(s, _worksheet.PositionID, Address);
		}
	}
	/// <summary>
	/// The named style
	/// </summary>
	public string StyleName
	{
		get
		{
			IsRangeValid("styling");
			int xfId;
			if (IsFullColumn())
			{
				xfId = GetColumnStyle(_fromCol);
			}
			else if (IsFullRow())
			{
				xfId = 0;
				if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref xfId))
				{
					xfId = GetColumnStyle(_fromCol);
				}
			}
			else
			{
				xfId = 0;
				if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref xfId))
				{
					if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref xfId))
					{
						xfId = GetColumnStyle(_fromCol);
					}
				}
			}

			return GetNamedStyleName(xfId);
		}
		set
		{
			_styleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
			if (IsFullColumn())    //Full column
			{
				SetStyleNameOnColumns(value);
			}
			else if (IsFullRow()) //FullRow
			{
				for (var r = _fromRow; r <= _toRow; r++)
				{
					_worksheet.Row(r)._styleName = value;
					_worksheet.Row(r).StyleID = _styleID;
				}
			}

			if (!(IsFullColumn() || IsFullRow())) //Cell specific
			{
				SetStyleIdOnEachCell();
			}
			else //Only set name on created cells. (uncreated cells is set on full row or full column).
			{
				var cells = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
				while (cells.Next())
				{
					_worksheet.SetStyleInner(cells.Row, cells.Column, _styleID);
				}
			}
		}
	}

	private bool IsFullColumn() => _fromRow == 1 && _toRow == ExcelPackage.MaxRows;

	private bool IsFullRow() => _fromCol == 1 && _toCol == ExcelPackage.MaxColumns;

	private string GetNamedStyleName(int xfId)
	{
		int nsID;
		if (xfId <= 0)
		{
			nsID = Style.Styles.CellXfs[0].XfId;
		}
		else
		{
			nsID = Style.Styles.CellXfs[xfId].XfId;
		}

		foreach (var ns in Style.Styles.NamedStyles)
		{
			if (ns.StyleXfId == nsID)
			{
				return ns.Name;
			}
		}

		return "";
	}

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
	private int GetColumnStyle(int col)
	{
		object c = null;
		if (_worksheet.ExistsValueInner(0, col, ref c))
		{
			return (c as ExcelColumn).StyleID;
		}
		else
		{
			var row = 0;
			if (_worksheet._values.PrevCell(ref row, ref col))
			{
				var column = _worksheet.GetValueInner(row, col) as ExcelColumn;
				if (column.ColumnMax >= col)
				{
					return _worksheet.GetStyleInner(row, col);
				}
			}
		}

		return 0;
	}
	/// <summary>
	/// The style ID. 
	/// It is not recomended to use this one. Use Named styles as an alternative.
	/// If you do, make sure that you use the Style.UpdateXml() method to update any new styles added to the workbook.
	/// </summary>
	public int StyleID
	{
		get
		{
			var s = 0;
			if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref s))
			{
				if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref s))
				{
					s = _worksheet.GetStyleInner(0, _fromCol);
				}
			}

			return s;
		}
		set
		{
			_changePropMethod(this, _setStyleIdDelegate, value);
		}
	}
}

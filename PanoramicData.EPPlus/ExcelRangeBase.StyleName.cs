namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
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
}

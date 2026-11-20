internal partial class CellStore<T>
{
	internal bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
	{
		if (ColumnCount == 0)
		{
			fromRow = fromCol = toRow = toCol = 0;
			return false;
		}
		else
		{
			fromCol = _columnIndex[0].Index;
			var fromIndex = 0;
			if (fromCol <= 0 && ColumnCount > 1)
			{
				fromCol = _columnIndex[1].Index;
				fromIndex = 1;
			}
			else if (ColumnCount == 1 && fromCol <= 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}

			var col = GetLastColumnPosition();
			toCol = _columnIndex[col].Index;
			if (toCol == 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}

			fromRow = toRow = 0;

			for (var c = fromIndex; c < ColumnCount; c++)
			{
				if (_columnIndex[c].PageCount == 0)
				{
					continue;
				}

				var first = GetFirstRow(c);
				var last = GetLastRow(c, first);
				if (first > 0 && (first < fromRow || fromRow == 0))
				{
					fromRow = first;
				}

				if (first > 0 && (last > toRow || toRow == 0))
				{
					toRow = last;
				}
			}

			if (fromRow <= 0 || toRow <= 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}
			else
			{
				return true;
			}
		}
	}

	private int GetLastColumnPosition()
	{
		var col = ColumnCount - 1;
		while (col > 0)
		{
			if (_columnIndex[col].PageCount == 0 || _columnIndex[col]._pages[0].RowCount > 1 || _columnIndex[col]._pages[0].Rows[0].Index > 0)
			{
				break;
			}

			col--;
		}

		return col;
	}

	private int GetFirstRow(int c)
	{
		int first;
		if (_columnIndex[c]._pages[0].RowCount > 0 && _columnIndex[c]._pages[0].Rows[0].Index > 0)
		{
			first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[0].Index;
		}
		else
		{
			if (_columnIndex[c]._pages[0].RowCount > 1)
			{
				first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[0].Rows[1].Index;
			}
			else if (_columnIndex[c].PageCount > 1)
			{
				first = _columnIndex[c]._pages[0].IndexOffset + _columnIndex[c]._pages[1].Rows[0].Index;
			}
			else
			{
				first = 0;
			}
		}

		return first;
	}

	private int GetLastRow(int c, int first)
	{
		int last;
		var lp = _columnIndex[c].PageCount - 1;
		while (_columnIndex[c]._pages[lp].RowCount == 0 && lp != 0)
		{
			lp--;
		}

		var p = _columnIndex[c]._pages[lp];
		if (p.RowCount > 0)
		{
			last = p.IndexOffset + p.Rows[p.RowCount - 1].Index;
		}
		else
		{
			last = first;
		}

		return last;
	}
}

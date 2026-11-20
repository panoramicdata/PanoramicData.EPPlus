internal partial class CellStore<T>
{
	///
	private static int DeleteCells(PageIndex page, int fromRow, int toRow, bool shift)
	{
		var fromPos = page.GetPosition(fromRow - page.IndexOffset);
		if (fromPos < 0)
		{
			fromPos = ~fromPos;
		}

		var maxRow = page.MaxIndex;
		var offset = toRow - page.IndexOffset;
		if (offset > PageSizeMax)
		{
			offset = PageSizeMax;
		}

		var toPos = page.GetPosition(offset);
		if (toPos < 0)
		{
			toPos = ~toPos;
		}

		if (fromPos <= toPos && fromPos < page.RowCount && page.GetIndex(fromPos) < toRow)
		{
			if (toRow > page.MaxIndex)
			{
				if (fromRow == page.MinIndex) //Delete entire page, late in the page delete method
				{
					return fromRow;
				}

				var r = page.MaxIndex;
				var deletedRow = page.RowCount - fromPos;
				page.RowCount -= deletedRow;
				return r + 1;
			}
			else
			{
				var rows = toRow - fromRow;
				if (shift)
				{
					UpdateRowIndex(page, toPos, rows);
				}

				Array.Copy(page.Rows, toPos, page.Rows, fromPos, page.RowCount - toPos);
				page.RowCount -= toPos - fromPos;

				return toRow;
			}
		}
		else if (shift)
		{
			UpdateRowIndex(page, toPos, toRow - fromRow);
		}

		return toRow < maxRow ? toRow : maxRow;
	}

	private void DeleteColumns(int fromCol, int columns, bool shift)
	{
		var fPos = GetPosition(fromCol);
		if (fPos < 0)
		{
			fPos = ~fPos;
		}

		var tPos = fPos;
		for (var c = fPos; c <= ColumnCount; c++)
		{
			tPos = c;
			if (tPos == ColumnCount || _columnIndex[c].Index >= fromCol + columns)
			{
				break;
			}
		}

		if (ColumnCount <= fPos)
		{
			return;
		}

		if (_columnIndex[fPos].Index >= fromCol && _columnIndex[fPos].Index <= fromCol + columns)
		{
			if (tPos < ColumnCount)
			{
				Array.Copy(_columnIndex, tPos, _columnIndex, fPos, ColumnCount - tPos);
			}

			ColumnCount -= tPos - fPos;
		}

		if (shift)
		{
			for (var c = fPos; c < ColumnCount; c++)
			{
				_columnIndex[c].Index -= (short)columns;
			}
		}
	}
}

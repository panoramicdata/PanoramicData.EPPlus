internal partial class CellStore<T>
{
	private void DeleteRowsInPage(ColumnIndex column, int pagePos, int fromRow, int rows, bool shift)
	{
		var endRow = fromRow + rows;
		var delEndRow = CellStore<T>.DeleteCells(column._pages[pagePos], fromRow, endRow, shift);
		if (shift && delEndRow != fromRow)
		{
			UpdatePageOffset(column, pagePos, delEndRow - fromRow);
		}

		if (endRow > delEndRow && pagePos < column.PageCount && column._pages[pagePos].MinIndex < endRow)
		{
			pagePos = delEndRow == fromRow ? pagePos : pagePos + 1;
			var rowsLeft = CellStore<T>.DeletePage(shift ? fromRow : delEndRow, endRow - delEndRow, column, pagePos, shift);
			if (rowsLeft > 0)
			{
				var fr = shift ? fromRow : endRow - rowsLeft;
				pagePos = column.GetPosition(fr);
				CellStore<T>.DeleteCells(column._pages[pagePos], fr, shift ? fr + rowsLeft : endRow, shift);
				if (shift)
				{
					UpdatePageOffset(column, pagePos, rowsLeft);
				}
			}
		}
	}

	private void UpdateOffsetsAfterDelete(ColumnIndex column, int pagePos, int fromRow, int rows, bool shift, int pageFromRow)
	{
		if (pagePos > 0 && column._pages[pagePos].IndexOffset > fromRow) //The row is on the page before.
		{
			var offset = fromRow + rows - 1 - ((pageFromRow - 1) << pageBits);
			var rowPos = column._pages[pagePos - 1].GetPosition(offset);
			if (rowPos > 0 && pagePos > 0)
			{
				if (shift)
				{
					UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow + rows - 1, -rows);
				}
			}
		}
		else
		{
			if (shift && pagePos + 1 < column.PageCount)
			{
				UpdateIndexOffset(column, pagePos + 1, 0, column._pages[pagePos + 1].MinIndex, -rows);
			}
		}
	}
}

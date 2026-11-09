using OfficeOpenXml;

internal partial class CellStore<T>
{
	internal void Clear(int fromRow, int fromCol, int rows, int columns) => Delete(fromRow, fromCol, rows, columns, false);
	internal void Delete(int fromRow, int fromCol, int rows, int columns) => Delete(fromRow, fromCol, rows, columns, true);
	internal void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
	{
		lock (_columnIndex)
		{
			if (columns > 0 && fromRow == 0 && rows >= ExcelPackage.MaxRows)
			{
				DeleteColumns(fromCol, columns, shift);
			}
			else
			{
				var toCol = fromCol + columns - 1;
				var pageFromRow = fromRow >> pageBits;
				for (var c = 0; c < ColumnCount; c++)
				{
					var column = _columnIndex[c];
					if (column.Index >= fromCol)
					{
						if (column.Index > toCol)
						{
							break;
						}

						DeleteRowsInColumn(column, fromRow, ref rows, shift, pageFromRow);
					}
				}
			}
		}
	}

	// rows is by ref: a shift on one column reduces the rows deleted from the following columns.
	private void DeleteRowsInColumn(ColumnIndex column, int fromRow, ref int rows, bool shift, int pageFromRow)
	{
		var pagePos = column.GetPosition(fromRow);
		if (pagePos < 0)
		{
			pagePos = ~pagePos;
		}

		if (pagePos < column.PageCount)
		{
			var page = column._pages[pagePos];
			if (shift && page.RowCount > 0 && page.MinIndex > fromRow && page.MaxIndex >= fromRow + rows)
			{
				var o = page.MinIndex - fromRow;
				if (o < rows)
				{
					rows -= o;
					page.Offset -= o;
					UpdatePageOffset(column, pagePos, o);
				}
				else
				{
					page.Offset -= rows;
					UpdatePageOffset(column, pagePos, rows);
					return;
				}
			}

			if (page.RowCount > 0 && page.MinIndex <= fromRow + rows - 1 && page.MaxIndex >= fromRow) //The row is inside the page
			{
				DeleteRowsInPage(column, pagePos, fromRow, rows, shift);
			}
			else
			{
				UpdateOffsetsAfterDelete(column, pagePos, fromRow, rows, shift, pageFromRow);
			}
		}
	}
}
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
									continue;
								}
							}

							if (page.RowCount > 0 && page.MinIndex <= fromRow + rows - 1 && page.MaxIndex >= fromRow) //The row is inside the page
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
							else if (pagePos > 0 && column._pages[pagePos].IndexOffset > fromRow) //The row is on the page before.
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
				}
			}
		}
	}
}

internal partial class CellStore<T>
{
	internal void Insert(int fromRow, int fromCol, int rows, int columns)
	{
		lock (_columnIndex)
		{

			if (columns > 0)
			{
				var col = GetPosition(fromCol);
				if (col < 0)
				{
					col = ~col;
				}

				for (var c = col; c < ColumnCount; c++)
				{
					_columnIndex[c].Index += (short)columns;
				}
			}
			else
			{
				var page = fromRow >> pageBits;
				for (var c = 0; c < ColumnCount; c++)
				{
					var column = _columnIndex[c];
					var pagePos = column.GetPosition(fromRow);
					if (pagePos >= 0)
					{
						if (IsWithinPage(fromRow, column, pagePos)) //The row is inside the page
						{
							var offset = fromRow - column._pages[pagePos].IndexOffset;
							var rowPos = column._pages[pagePos].GetPosition(offset);
							if (rowPos < 0)
							{
								rowPos = ~rowPos;
							}

							UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
						}
						else if (pagePos > 0 && IsWithinPage(fromRow, column, pagePos - 1)) //The row is inside the previous page
						{
							var offset = fromRow - ((page - 1) << pageBits);
							var rowPos = column._pages[pagePos - 1].GetPosition(offset);
							if (rowPos > 0 && pagePos > 0)
							{
								UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow, rows);
							}
						}
						else if (column.PageCount >= pagePos + 1)
						{
							var offset = fromRow - column._pages[pagePos].IndexOffset;
							var rowPos = column._pages[pagePos].GetPosition(offset);
							if (rowPos < 0)
							{
								rowPos = ~rowPos;
							}

							if (column._pages[pagePos].RowCount > rowPos)
							{
								UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
							}
							else
							{
								UpdateIndexOffset(column, pagePos + 1, 0, fromRow, rows);
							}
						}
					}
					else
					{
						UpdateIndexOffset(column, ~pagePos, 0, fromRow, rows);
					}
				}
			}
		}
	}

	private static bool IsWithinPage(int row, ColumnIndex column, int pagePos) => (row >= column._pages[pagePos].MinIndex && row <= column._pages[pagePos].MaxIndex);

	private void SplitPageInsert(ColumnIndex column, int pagePos, int rowPos, int rows, int size, int addPages)
	{
		var page = column._pages[pagePos];

		var rStart = -1;
		for (var r = rowPos; r < page.RowCount; r++)
		{
			if (page.IndexExpanded - (page.Rows[r].Index + rows) > PageSize)
			{
				rStart = r;
				break;
			}
			else
			{
				page.Rows[r].Index += (short)rows;
			}
		}

		var rc = page.RowCount - rStart;
		page.RowCount = rStart;
		if (rc > 0)
		{
			//Copy to a new page
			var newPage = CellStore<T>.CopyNew(page, rStart, rc);
			var ix = (short)(page.Index + addPages);
			var offset = page.IndexOffset + rows - (ix * PageSize);
			if (offset > PageSize)
			{
				ix += (short)(offset / PageSize);
				offset %= PageSize;
			}

			newPage.Index = ix;
			newPage.Offset = offset;
			CellStore<T>.AddPage(column, pagePos + 1, newPage);
		}

		//Copy from next Row
	}

	private static void CopyMergePage(PageIndex page, int rowPos, int rows, int size, PageIndex ToPage)
	{
		var newRows = new IndexItem[GetSize(ToPage.RowCount + size)];
		page.RowCount -= size;
		Array.Copy(page.Rows, rowPos, newRows, 0, size);
		for (var r = 0; r < size; r++)
		{
			newRows[r].Index += (short)(page.IndexOffset + rows - ToPage.IndexOffset);
		}

		Array.Copy(ToPage.Rows, 0, newRows, size, ToPage.RowCount);
		ToPage.Rows = newRows;
		ToPage.RowCount += size;
	}
	private static void MergePage(ColumnIndex column, int pagePos)
	{
		var Page1 = column._pages[pagePos];
		var Page2 = column._pages[pagePos + 1];

		var newPage = new PageIndex(Page1, 0, Page1.RowCount + Page2.RowCount)
		{
			RowCount = Page1.RowCount + Page2.RowCount
		};
		Array.Copy(Page1.Rows, 0, newPage.Rows, 0, Page1.RowCount);
		Array.Copy(Page2.Rows, 0, newPage.Rows, Page1.RowCount, Page2.RowCount);
		for (var r = Page1.RowCount; r < newPage.RowCount; r++)
		{
			newPage.Rows[r].Index += (short)(Page2.IndexOffset - Page1.IndexOffset);
		}

		column._pages[pagePos] = newPage;
		column.PageCount--;

		if (column.PageCount > (pagePos + 1))
		{
			Array.Copy(column._pages, pagePos + 2, column._pages, pagePos + 1, column.PageCount - (pagePos + 1));
			for (var p = pagePos + 1; p < column.PageCount; p++)
			{
				column._pages[p].Index--;
				column._pages[p].Offset += PageSize;
			}
		}
	}

	private static PageIndex CopyNew(PageIndex pageFrom, int rowPos, int size)
	{
		var newRows = new IndexItem[GetSize(size)];
		Array.Copy(pageFrom.Rows, rowPos, newRows, 0, size);
		return new PageIndex(newRows, size);
	}

	internal static int GetSize(int size)
	{
		var newSize = 256;
		while (newSize < size)
		{
			newSize <<= 1;
		}

		return newSize;
	}
}

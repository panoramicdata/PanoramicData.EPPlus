internal partial class CellStore<T>
{
	private void UpdateIndexOffset(ColumnIndex column, int pagePos, int rowPos, int row, int rows)
	{
		if (pagePos >= column.PageCount)
		{
			return;    //A page after last cell.
		}

		var page = column._pages[pagePos];
		if (rows > PageSize)
		{
			var addPages = (short)(rows >> pageBits);
			var offset = +(rows - (PageSize * addPages));
			for (var p = pagePos + 1; p < column.PageCount; p++)
			{
				if (column._pages[p].Offset + offset > PageSize)
				{
					column._pages[p].Index += (short)(addPages + 1);
					column._pages[p].Offset += offset - PageSize;
				}
				else
				{
					column._pages[p].Index += addPages;
					column._pages[p].Offset += offset;
				}
			}

			var size = page.RowCount - rowPos;
			if (page.RowCount > rowPos)
			{
				if (column.PageCount - 1 == pagePos) //No page after, create a new one.
				{
					//Copy rows to next page.
					var newPage = CellStore<T>.CopyNew(page, rowPos, size);
					newPage.Index = (short)((row + rows) >> pageBits);
					newPage.Offset = row + rows - (newPage.Index * PageSize) - newPage.Rows[0].Index;
					if (newPage.Offset > PageSize)
					{
						newPage.Index++;
						newPage.Offset -= PageSize;
					}

					CellStore<T>.AddPage(column, pagePos + 1, newPage);
					page.RowCount = rowPos;
				}
				else
				{
					if (column._pages[pagePos + 1].RowCount + size > PageSizeMax) //Split Page
					{
						SplitPageInsert(column, pagePos, rowPos, rows, addPages);
					}
					else //Copy Page.
					{
						CellStore<T>.CopyMergePage(page, rowPos, rows, size, column._pages[pagePos + 1]);
					}
				}
			}
		}
		else
		{
			//Add to Pages.
			for (var r = rowPos; r < page.RowCount; r++)
			{
				page.Rows[r].Index += (short)rows;
			}

			if (page.Offset + page.Rows[page.RowCount - 1].Index >= PageSizeMax)   //Can not be larger than the max size of the page.
			{
				CellStore<T>.AdjustIndex(column, pagePos);
				if (page.Offset + page.Rows[page.RowCount - 1].Index >= PageSizeMax)
				{
					pagePos = CellStore<T>.SplitPage(column, pagePos);
				}
				//TODO: MoveRows to next page.
			}

			for (var p = pagePos + 1; p < column.PageCount; p++)
			{
				if (column._pages[p].Offset + rows < PageSize)
				{
					column._pages[p].Offset += rows;
				}
				else
				{
					column._pages[p].Index++;
					column._pages[p].Offset = (column._pages[p].Offset + rows) % PageSize;
				}
			}
		}
	}
	private static void AddPage(ColumnIndex column, int pos, short index)
	{
		CellStore<T>.AddPage(column, pos);
		column._pages[pos] = new PageIndex { Index = index };
		if (pos > 0)
		{
			var pp = column._pages[pos - 1];
			if (pp.RowCount > 0 && pp.Rows[pp.RowCount - 1].Index > PageSize)
			{
				column._pages[pos].Offset = pp.Rows[pp.RowCount - 1].Index - PageSize;
			}
		}
	}
	/// <summary>
	/// Add a new page to the collection
	/// </summary>
	/// <param name="column">The column</param>
	/// <param name="pos">Position</param>
	/// <param name="page">The new page object to add</param>
	private static void AddPage(ColumnIndex column, int pos, PageIndex page)
	{
		CellStore<T>.AddPage(column, pos);
		column._pages[pos] = page;
	}
	/// <summary>
	/// Add a new page to the collection
	/// </summary>
	/// <param name="column">The column</param>
	/// <param name="pos">Position</param>
	private static void AddPage(ColumnIndex column, int pos)
	{
		if (column.PageCount == column._pages.Length)
		{
			var pageTmp = new PageIndex[column._pages.Length * 2];
			Array.Copy(column._pages, 0, pageTmp, 0, column.PageCount);
			column._pages = pageTmp;
		}

		if (pos < column.PageCount)
		{
			Array.Copy(column._pages, pos, column._pages, pos + 1, column.PageCount - pos);
		}

		column.PageCount++;
	}
	private void AddColumn(int pos, int Column)
	{
		if (ColumnCount == _columnIndex.Length)
		{
			var colTmp = new ColumnIndex[_columnIndex.Length * 2];
			Array.Copy(_columnIndex, 0, colTmp, 0, ColumnCount);
			_columnIndex = colTmp;
		}

		if (pos < ColumnCount)
		{
			Array.Copy(_columnIndex, pos, _columnIndex, pos + 1, ColumnCount - pos);
		}

		_columnIndex[pos] = new ColumnIndex { Index = (short)Column };
		ColumnCount++;
	}
}

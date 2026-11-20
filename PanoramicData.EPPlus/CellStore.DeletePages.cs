internal partial class CellStore<T>
{
	private void UpdatePageOffset(ColumnIndex column, int pagePos, int rows)
	{
		//Update Pageoffset

		if (++pagePos < column.PageCount)
		{
			for (var p = pagePos; p < column.PageCount; p++)
			{
				if (column._pages[p].Offset - rows <= -PageSize)
				{
					column._pages[p].Index--;
					column._pages[p].Offset -= rows - PageSize;
				}
				else
				{
					column._pages[p].Offset -= rows;
				}
			}

			if (Math.Abs(column._pages[pagePos].Offset) > PageSize ||
				Math.Abs(column._pages[pagePos].Rows[column._pages[pagePos].RowCount - 1].Index) > PageSizeMax) //Split or Merge???
			{
				CellStore<T>.ResetPageOffset(column, pagePos);
			}
		}
	}

	private static void ResetPageOffset(ColumnIndex column, int pagePos)
	{
		var fromPage = column._pages[pagePos];
		PageIndex toPage;
		short pageAdd = 0;
		if (fromPage.Offset < -PageSize)
		{
			toPage = column._pages[pagePos - 1];
			pageAdd = -1;
			if (fromPage.Index - 1 == toPage.Index)
			{
				if (fromPage.IndexOffset + fromPage.Rows[fromPage.RowCount - 1].Index -
					toPage.IndexOffset + toPage.Rows[0].Index <= PageSizeMax)
				{
					CellStore<T>.MergePage(column, pagePos - 1);
				}
			}
			else //No page after 
			{
				fromPage.Index -= pageAdd;
				fromPage.Offset += PageSize;
			}
		}
		else if (fromPage.Offset > PageSize)
		{
			toPage = column._pages[pagePos + 1];
			pageAdd = 1;
			if (fromPage.Index + 1 != toPage.Index)
			{
				fromPage.Index += pageAdd;
				fromPage.Offset += PageSize;
			}
		}
	}

	private static int DeletePage(int fromRow, int rows, ColumnIndex column, int pagePos, bool shift)
	{
		var page = column._pages[pagePos];
		var startRows = rows;
		while (page != null && page.MinIndex >= fromRow && ((shift && page.MaxIndex < fromRow + rows) || (!shift && page.MaxIndex < fromRow + startRows)))
		{
			//Delete entire page.
			var delSize = page.MaxIndex - page.MinIndex + 1;
			rows -= delSize;
			Array.Copy(column._pages, pagePos + 1, column._pages, pagePos, column.PageCount - pagePos - 1);
			column.PageCount--;
			if (column.PageCount == 0)
			{
				return 0;
			}

			if (shift)
			{
				for (var i = pagePos; i < column.PageCount; i++)
				{
					column._pages[i].Offset -= delSize;
					if (column._pages[i].Offset <= -PageSize)
					{
						column._pages[i].Index--;
						column._pages[i].Offset += PageSize;
					}
				}
			}

			if (column.PageCount > pagePos)
			{
				page = column._pages[pagePos];
			}
			else
			{
				//No more pages, return 0
				return 0;
			}
		}

		return rows;
	}

	private static void UpdateRowIndex(PageIndex page, int toPos, int rows)
	{
		for (var r = toPos; r < page.RowCount; r++)
		{
			page.Rows[r].Index -= (short)rows;
		}
	}
}

internal partial class CellStore<T>
{
	private void AddCell(ColumnIndex columnIndex, int pagePos, int pos, short ix, T value)
	{
		var pageItem = columnIndex._pages[pagePos];
		if (pageItem.RowCount == pageItem.Rows.Length)
		{
			if (pageItem.RowCount == PageSizeMax) //Max size-->Split
			{
				pagePos = CellStore<T>.SplitPage(columnIndex, pagePos);
				if (columnIndex._pages[pagePos - 1].RowCount > pos)
				{
					pagePos--;
				}
				else
				{
					pos -= columnIndex._pages[pagePos - 1].RowCount;
				}

				pageItem = columnIndex._pages[pagePos];
			}
			else //Expand to double size.
			{
				var rowsTmp = new IndexItem[pageItem.Rows.Length << 1];
				Array.Copy(pageItem.Rows, 0, rowsTmp, 0, pageItem.RowCount);
				pageItem.Rows = rowsTmp;
			}
		}

		if (pos < pageItem.RowCount)
		{
			Array.Copy(pageItem.Rows, pos, pageItem.Rows, pos + 1, pageItem.RowCount - pos);
		}

		pageItem.Rows[pos] = new IndexItem { Index = ix, IndexPointer = _values.Count };
		_values.Add(value);
		pageItem.RowCount++;
	}

	private static int SplitPage(ColumnIndex columnIndex, int pagePos)
	{
		var page = columnIndex._pages[pagePos];
		if (page.Offset != 0)
		{
			var offset = page.Offset;
			page.Offset = 0;
			for (var r = 0; r < page.RowCount; r++)
			{
				page.Rows[r].Index -= (short)offset;
			}
		}
		//Find Split pos
		var splitPos = 0;
		for (var r = 0; r < page.RowCount; r++)
		{
			if (page.Rows[r].Index > PageSize)
			{
				splitPos = r;
				break;
			}
		}

		var newPage = new PageIndex(page, 0, splitPos);
		var nextPage = new PageIndex(page, splitPos, page.RowCount - splitPos, (short)(page.Index + 1), page.Offset);

		for (var r = 0; r < nextPage.RowCount; r++)
		{
			nextPage.Rows[r].Index = (short)(nextPage.Rows[r].Index - PageSize);
		}

		columnIndex._pages[pagePos] = newPage;
		if (columnIndex.PageCount + 1 > columnIndex._pages.Length)
		{
			var pageTmp = new PageIndex[columnIndex._pages.Length << 1];
			Array.Copy(columnIndex._pages, 0, pageTmp, 0, columnIndex.PageCount);
			columnIndex._pages = pageTmp;
		}

		Array.Copy(columnIndex._pages, pagePos + 1, columnIndex._pages, pagePos + 2, columnIndex.PageCount - pagePos - 1);
		columnIndex._pages[pagePos + 1] = nextPage;
		columnIndex.PageCount++;
		return pagePos + 1;
	}

	private static void AdjustIndex(ColumnIndex columnIndex, int pagePos)
	{
		var page = columnIndex._pages[pagePos];
		//First Adjust indexes
		if (page.Offset + page.Rows[0].Index >= PageSize ||
			page.Offset >= PageSize ||
			page.Rows[0].Index >= PageSize)
		{
			page.Index++;
			page.Offset -= PageSize;
		}
		else if (page.Offset + page.Rows[0].Index <= -PageSize ||
				 page.Offset <= -PageSize ||
				 page.Rows[0].Index <= -PageSize)
		{
			page.Index--;
			page.Offset += PageSize;
		}
	}

	private static void AddPageRowOffset(PageIndex page, short offset)
	{
		for (var r = 0; r < page.RowCount; r++)
		{
			page.Rows[r].Index += offset;
		}
	}
}

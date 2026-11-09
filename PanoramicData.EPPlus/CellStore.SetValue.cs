using System.Diagnostics.CodeAnalysis;

internal partial class CellStore<T>
{
	private int GetOrAddPagePosition(int col, int row, short page)
	{
		var pos = _columnIndex[col].GetPosition(row);
		if (pos < 0)
		{
			pos = ~pos;
			if (pos - 1 < 0 || _columnIndex[col]._pages[pos - 1].IndexOffset + PageSize - 1 < row)
			{
				CellStore<T>.AddPage(_columnIndex[col], pos, page);
			}
			else
			{
				pos--;
			}
		}

		if (pos >= _columnIndex[col].PageCount)
		{
			CellStore<T>.AddPage(_columnIndex[col], pos, page);
		}

		return pos;
	}

	private int AddCellInNewColumn(int col, int column, int row, short page, T? value)
	{
		AddColumn(col, column);
		CellStore<T>.AddPage(_columnIndex[col], 0, page);
		var ix = (short)(row - (page << pageBits));
		AddCell(_columnIndex[col], 0, 0, ix, value);
		return col;
	}

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal void SetValue(int Row, int Column, T? Value)
	{
		lock (_columnIndex)
		{
			var col = GetPosition(Column);
			var page = (short)(Row >> pageBits);
			if (col >= 0)
			{
				var pos = GetOrAddPagePosition(col, Row, page);
				var pageItem = _columnIndex[col]._pages[pos];
				if (!(pageItem.MinIndex <= Row && pageItem.MaxIndex >= Row) && pageItem.IndexExpanded > Row)   //TODO: Fix issue
				{
					pos--;
					if (pos < 0)
					{
						throw new Exception("Unexpected error when setting value");
					}

					pageItem = _columnIndex[col]._pages[pos];
				}

				var ix = (short)(Row - ((pageItem.Index << pageBits) + pageItem.Offset));
				_searchItem.Index = ix;
				var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
				if (cellPos < 0)
				{
					cellPos = ~cellPos;
					AddCell(_columnIndex[col], pos, cellPos, ix, Value);
				}
				else
				{
					_values[pageItem.Rows[cellPos].IndexPointer] = Value;
				}
			}
			else //Column does not exist
			{
				AddCellInNewColumn(~col, Column, Row, page, Value);
			}
		}
	}

	internal delegate void SetRangeValueDelegate(List<T> list, int index, int row, int column, object value);
	/// <summary>
	/// Set Value for Range
	/// </summary>
	/// <param name="fromRow"></param>
	/// <param name="fromColumn"></param>
	/// <param name="toRow"></param>
	/// <param name="toColumn"></param>
	/// <param name="Updater"></param>
	/// <param name="Value"></param>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal void SetRangeValueSpecial(int fromRow, int fromColumn, int toRow, int toColumn, SetRangeValueDelegate Updater, object Value)
	{
		lock (_columnIndex)
		{
			// split row to page groups (pageIndex to RowNo List)
			Dictionary<short, List<int>> pages = [];
			for (var rowIx = fromRow; rowIx <= toRow; rowIx++)
			{
				var pageIx = (short)(rowIx >> pageBits);
				if (!pages.ContainsKey(pageIx))
				{
					pages.Add(pageIx, []);
				}

				pages[pageIx].Add(rowIx);
			}

			for (var colIx = fromColumn; colIx <= toColumn; colIx++)
			{
				var col = GetPosition(colIx);

				foreach (var pair in pages)
				{
					var page = pair.Key;
					foreach (var rowIx in pair.Value)
					{
						if (col >= 0)
						{
							var pos = GetOrAddPagePosition(col, rowIx, page);
							var pageItem = _columnIndex[col]._pages[pos];
							if (pageItem.IndexOffset > rowIx)
							{
								pos--;
								page--;
								if (pos < 0)
								{
									throw new Exception("Unexpected error when setting value");
								}

								pageItem = _columnIndex[col]._pages[pos];
							}

							var ix = (short)(rowIx - ((pageItem.Index << pageBits) + pageItem.Offset));
							_searchItem.Index = ix;
							var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
							if (cellPos < 0)
							{
								cellPos = ~cellPos;
								AddCell(_columnIndex[col], pos, cellPos, ix, default);
								Updater(_values, pageItem.Rows[cellPos].IndexPointer, rowIx, colIx, Value);
							}
							else
							{
								Updater(_values, pageItem.Rows[cellPos].IndexPointer, rowIx, colIx, Value);
							}
						}
						else //Column does not exist
						{
							col = AddCellInNewColumn(~col, colIx, rowIx, page, default);
							Updater(_values, _columnIndex[col]._pages[0].Rows[0].IndexPointer, rowIx, colIx, Value);
						}
					}
				}
			}
		}
	}

	internal delegate void SetValueDelegate(List<T> list, int index, object value);
	// Set object's property atomically
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal void SetValueSpecial(int Row, int Column, SetValueDelegate Updater, object? Value)
	{
		lock (_columnIndex)
		{
			var col = GetPosition(Column);
			var page = (short)(Row >> pageBits);
			if (col >= 0)
			{
				var pos = GetOrAddPagePosition(col, Row, page);
				var pageItem = _columnIndex[col]._pages[pos];
				if (pageItem.IndexOffset > Row)
				{
					pos--;
					if (pos < 0)
					{
						throw new Exception("Unexpected error when setting value");
					}

					pageItem = _columnIndex[col]._pages[pos];
				}

				var ix = (short)(Row - ((pageItem.Index << pageBits) + pageItem.Offset));
				_searchItem.Index = ix;
				var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
				if (cellPos < 0)
				{
					cellPos = ~cellPos;
					AddCell(_columnIndex[col], pos, cellPos, ix, default);
					Updater(_values, pageItem.Rows[cellPos].IndexPointer, Value);
				}
				else
				{
					Updater(_values, pageItem.Rows[cellPos].IndexPointer, Value);
				}
			}
			else //Column does not exist
			{
				col = AddCellInNewColumn(~col, Column, Row, page, default);
				Updater(_values, _columnIndex[col]._pages[0].Rows[0].IndexPointer, Value);
			}
		}
	}
}

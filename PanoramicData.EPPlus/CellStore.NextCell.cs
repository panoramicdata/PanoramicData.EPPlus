using OfficeOpenXml;

internal partial class CellStore<T>
{
	internal bool NextCell(ref int row, ref int col) => NextCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
	internal bool NextCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
	{
		if (!ClampColumnPositions(minColPos, ref maxColPos))
		{
			return false;
		}

		var c = GetPosition(col);
		if (c >= 0)
		{
			if (c > maxColPos)
			{
				if (col <= minColPos)
				{
					return false;
				}

				col = minColPos;
				return NextCell(ref row, ref col);
			}
			else
			{
				return GetNextCellInColumn(ref row, ref col, c, minColPos, maxRow, maxColPos);
			}
		}
		else
		{
			c = ~c;
			if (c >= ColumnCount)
			{
				c = ColumnCount - 1;
			}

			if (col > _columnIndex[c].Index)
			{
				if (col <= minColPos)
				{
					return false;
				}

				col = minColPos;
				return NextCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			else
			{
				return GetNextCellInColumn(ref row, ref col, c, minColPos, maxRow, maxColPos);
			}
		}
	}
	internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos)
	{
		if (ColumnCount == 0)
		{
			return false;
		}
		else
		{
			if (++colPos < ColumnCount && colPos <= endColPos)
			{
				var r = _columnIndex[colPos].GetNextRow(row);
				if (r == row) //Exists next Row
				{
					return true;
				}
				else
				{
					int minRow, minCol;
					if (r > row)
					{
						minRow = r;
						minCol = colPos;
					}
					else
					{
						minRow = int.MaxValue;
						minCol = 0;
					}

					if (FindRowInColumnsForward(row, colPos + 1, endColPos, ref colPos, ref minRow, ref minCol))
					{
						return true;
					}

					if (row < endRow)
					{
						row++;
						if (FindRowInEarlierColumns(row, startColPos, endRow, ref colPos, ref minRow, ref minCol))
						{
							return true;
						}
					}

					if (minRow == int.MaxValue || minRow > endRow)
					{
						return false;
					}
					else
					{
						row = minRow;
						colPos = minCol;
						return true;
					}
				}
			}
			else
			{
				if (colPos <= startColPos || row >= endRow)
				{
					return false;
				}

				colPos = startColPos - 1;
				row++;
				return GetNextCell(ref row, ref colPos, startColPos, endRow, endColPos);
			}
		}
	}

	private bool GetNextCellInColumn(ref int row, ref int col, int c, int minColPos, int maxRow, int maxColPos)
	{
		var r = GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
		col = _columnIndex[c].Index;
		return r;
	}
}

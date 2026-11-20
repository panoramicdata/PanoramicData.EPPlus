internal partial class CellStore<T>
{
	private bool FindRowInColumnsForward(int row, int c, int endColPos, ref int colPos, ref int minRow, ref int minCol)
	{
		while (c < ColumnCount && c <= endColPos)
		{
			var r = _columnIndex[c].GetNextRow(row);
			if (r == row) //Exists next Row
			{
				colPos = c;
				return true;
			}

			if (r > row && r < minRow)
			{
				minRow = r;
				minCol = c;
			}

			c++;
		}

		return false;
	}

	private bool FindRowInEarlierColumns(int row, int c, int endRow, ref int colPos, ref int minRow, ref int minCol)
	{
		while (c < colPos)
		{
			var r = _columnIndex[c].GetNextRow(row);
			if (r == row) //Exists next Row
			{
				colPos = c;
				return true;
			}

			if (r > row && (r < minRow || (r == minRow && c < minCol)) && r <= endRow)
			{
				minRow = r;
				minCol = c;
			}

			c++;
		}

		return false;
	}

	private bool ClampColumnPositions(int minColPos, ref int maxColPos)
	{
		if (minColPos >= ColumnCount)
		{
			return false;
		}

		if (maxColPos >= ColumnCount)
		{
			maxColPos = ColumnCount - 1;
		}

		return true;
	}
}

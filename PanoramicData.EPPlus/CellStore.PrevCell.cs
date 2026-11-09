using OfficeOpenXml;
using System.Diagnostics.CodeAnalysis;

internal partial class CellStore<T>
{
	private bool GetPrevCellInColumn(ref int row, ref int col, int c, int minRow, int minColPos, int maxColPos)
	{
		var ret = GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
		if (ret)
		{
			col = _columnIndex[c].Index;
		}

		return ret;
	}

	internal bool PrevCell(ref int row, ref int col) => PrevCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
	internal bool PrevCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
	{
		if (!ClampColumnPositions(minColPos, ref maxColPos))
		{
			return false;
		}

		var c = GetPosition(col);
		if (c >= 0)
		{
			if (c == 0)
			{
				if (col >= maxColPos)
				{
					return false;
				}

				if (row == minRow)
				{
					return false;
				}

				row--;
				col = maxColPos;
				return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			else
			{
				return GetPrevCellInColumn(ref row, ref col, c, minRow, minColPos, maxColPos);
			}
		}
		else
		{
			c = ~c;
			if (c == 0)
			{
				if (col >= maxColPos || row <= 0)
				{
					return false;
				}

				col = maxColPos;
				row--;
				return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			else
			{
				return GetPrevCellInColumn(ref row, ref col, c, minRow, minColPos, maxColPos);
			}
		}
	}
	[SuppressMessage("Major Bug", "S2234:Arguments should be passed in the same order as the method parameters", Justification = "Upstream EPPlus behaviour; reordering the recursive call would change enumeration results.")]
	internal bool GetPrevCell(ref int row, ref int colPos, int startRow, int startColPos, int endColPos)
	{
		if (ColumnCount == 0)
		{
			return false;
		}
		else
		{
			if (--colPos >= startColPos)
			{
				var r = _columnIndex[colPos].GetNextRow(row);
				if (r == row) //Exists next Row
				{
					return true;
				}
				else
				{
					int minRow, minCol;
					if (r > row && r >= startRow)
					{
						minRow = r;
						minCol = colPos;
					}
					else
					{
						minRow = int.MaxValue;
						minCol = 0;
					}

					var c = colPos - 1;
					if (c >= startColPos && FindRowInColumnsBackward(row, startRow, c, startColPos, ref colPos, ref minRow, ref minCol))
					{
						return true;
					}

					if (row > startRow)
					{
						c = endColPos;
						row--;
						if (FindRowInColumnsBackward(row, startRow, c, colPos + 1, ref colPos, ref minRow, ref minCol))
						{
							return true;
						}
					}

					if (minRow == int.MaxValue || startRow < minRow)
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
				colPos = ColumnCount;
				row--;
				if (row < startRow)
				{
					Reset();
					return false;
				}
				else
				{
					return GetPrevCell(ref colPos, ref row, startRow, startColPos, endColPos);
				}
			}
		}
	}

	private bool FindRowInColumnsBackward(int row, int startRow, int c, int lowestColPos, ref int colPos, ref int minRow, ref int minCol)
	{
		while (c >= lowestColPos)
		{
			var r = _columnIndex[c].GetNextRow(row);
			if (r == row) //Exists next Row
			{
				colPos = c;
				return true;
			}

			if (r > row && r < minRow && r >= startRow)
			{
				minRow = r;
				minCol = c;
			}

			c--;
		}

		return false;
	}
}
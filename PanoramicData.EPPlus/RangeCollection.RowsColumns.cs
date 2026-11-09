using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

internal partial class RangeCollection
{
	/// <summary>
	/// Insert a number of rows in the collecion but dont update the cell only the index
	/// </summary>
	/// <param name="rowID"></param>
	/// <param name="rows"></param>
	/// <returns>Index of first rangeItem</returns>
	internal int InsertRowsUpdateIndex(ulong rowID, int rows)
	{
		var index = IndexOf(rowID);
		if (index < 0)
		{
			index = ~index; //No match found invert to get start cell
		}

		var rowAdd = (((ulong)rows) << 29);
		for (var i = index; i < _cells.Count; i++)
		{
			_cellIndex[i].RangeID += rowAdd;
		}

		return index;
	}
	/// <summary>
	/// Insert a number of rows in the collecion
	/// </summary>
	/// <param name="rowID"></param>
	/// <param name="rows"></param>
	/// <returns>Index of first rangeItem</returns>
	internal int InsertRows(ulong rowID, int rows)
	{
		var index = IndexOf(rowID);
		if (index < 0)
		{
			index = ~index; //No match found invert to get start cell
		}

		var rowAdd = (((ulong)rows) << 29);
		for (var i = index; i < _cells.Count; i++)
		{
			_cellIndex[i].RangeID += rowAdd;
			_cells[_cellIndex[i].ListPointer].RangeID += rowAdd;
		}

		return index;
	}
	/// <summary>
	/// Delete rows from the collecion
	/// </summary>
	/// <param name="rowID"></param>
	/// <param name="rows"></param>
	/// <param name="updateCells">Update range id's on cells</param>
	internal int DeleteRows(ulong rowID, int rows, bool updateCells)
	{
		var rowAdd = (((ulong)rows) << 29);
		var index = IndexOf(rowID);
		if (index < 0)
		{
			index = ~index; //No match found invert to get start cell
		}

		if (index >= _cells.Count || _cellIndex[index] == null)
		{
			return -1;   //No row above this row
		}

		while (index < _cells.Count && _cellIndex[index].RangeID < rowID + rowAdd)
		{
			Delete(_cellIndex[index].RangeID);
		}

		var updIndex = IndexOf(rowID + rowAdd);
		if (updIndex < 0)
		{
			updIndex = ~updIndex; //No match found invert to get start cell
		}

		for (var i = updIndex; i < _cells.Count; i++)
		{
			_cellIndex[i].RangeID -= rowAdd;                        //Change the index
			if (updateCells)
			{
				_cells[_cellIndex[i].ListPointer].RangeID -= rowAdd;    //Change the cell/row or column object
			}
		}

		return index;
	}
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal static void InsertColumn(ulong ColumnID, int columns) => throw new Exception("Working on it...");
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal static void DeleteColumn(ulong ColumnID, int columns) => throw new Exception("Working on it...");
}

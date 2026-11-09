/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
     * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Added       		        2012-11-25
 *******************************************************************************/
using OfficeOpenXml;

/// <summary>
/// This is the store for all Rows, Columns and Cells.
/// It is a Dictionary implementation that allows you to change the Key (the RowID, ColumnID or CellID )
/// </summary>
internal partial class CellStore<T> : IDisposable// : IEnumerable<ulong>, IEnumerator<ulong>
{
	/**** Size constants ****/
	internal const int pageBits = 10;   //13bits=8192  Note: Maximum is 13 bits since short is used (PageMax=16K)
	internal const int PageSize = 1 << pageBits;
	internal const int PageSizeMin = 1 << 10;
	internal const int PageSizeMax = PageSize << 1; //Double page size
	internal const int ColSizeMin = 32;
	internal const int PagesPerColumnMin = 32;

	List<T>? _values = [];
	internal ColumnIndex[]? _columnIndex;
	internal IndexBase _searchIx = new();
	internal IndexItem _searchItem = new();
	internal int ColumnCount;
	public CellStore()
	{
		_columnIndex = new ColumnIndex[ColSizeMin];
	}
	~CellStore()
	{
		if (_values != null)
		{
			_values.Clear();
			_values = null;
		}

		_columnIndex = null;
	}
	internal int GetPosition(int Column)
	{
		if (Column < ColumnCount && _columnIndex[Column].Index == Column)      //Check if th column is lesser than
		{
			return Column;
		}
		else
		{
			_searchIx.Index = (short)Column;
			return Array.BinarySearch(_columnIndex, 0, ColumnCount, _searchIx);
		}
	}
	internal CellStore<T> Clone()
	{
		int row, col;
		var ret = new CellStore<T>();
		for (var c = 0; c < ColumnCount; c++)
		{
			col = _columnIndex[c].Index;
			for (var p = 0; p < _columnIndex[c].PageCount; p++)
			{
				for (var r = 0; r < _columnIndex[c]._pages[p].RowCount; r++)
				{
					row = _columnIndex[c]._pages[p].IndexOffset + _columnIndex[c]._pages[p].Rows[r].Index;
					ret.SetValue(row, col, _values[_columnIndex[c]._pages[p].Rows[r].IndexPointer]);
				}
			}
		}

		return ret;
	}
	internal int Count
	{
		get
		{
			var count = 0;
			for (var c = 0; c < ColumnCount; c++)
			{
				for (var p = 0; p < _columnIndex[c].PageCount; p++)
				{
					count += _columnIndex[c]._pages[p].RowCount;
				}
			}

			return count;
		}
	}
	internal int FindNext(int Column)
	{
		var c = GetPosition(Column);
		return c < 0 ? ~c : c;
	}
	internal T GetValue(int Row, int Column)
	{
		var i = GetPointer(Row, Column);
		return i >= 0 ? _values[i] : default;
	}
	int GetPointer(int Row, int Column)
	{
		var col = GetPosition(Column);
		if (col >= 0)
		{
			var pos = _columnIndex[col].GetPosition(Row);
			if (pos >= 0 && pos < _columnIndex[col].PageCount)
			{
				var pageItem = _columnIndex[col]._pages[pos];
				if (pageItem.MinIndex > Row)
				{
					pos--;
					if (pos < 0)
					{
						return -1;
					}
					else
					{
						pageItem = _columnIndex[col]._pages[pos];
					}
				}

				var ix = (short)(Row - pageItem.IndexOffset);
				_searchItem.Index = ix;
				var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
				return cellPos >= 0 ? pageItem.Rows[cellPos].IndexPointer : -1;
			}
			else //Page does not exist
			{
				return -1;
			}
		}
		else //Column does not exist
		{
			return -1;
		}
	}
	internal bool Exists(int Row, int Column) => GetPointer(Row, Column) >= 0;
	internal bool Exists(int Row, int Column, ref T value)
	{
		var p = GetPointer(Row, Column);
		if (p >= 0)
		{
			value = _values[p];
			return true;
		}
		else
		{
			return false;
		}
	}
	int _colPos = -1, _row;
	public ulong Current => ((ulong)_row << 32) | (ulong)(long)(_columnIndex[_colPos].Index);

	public void Dispose()
	{
		_values?.Clear();
		for (var c = 0; c < ColumnCount; c++)
		{
			if (_columnIndex[c] != null)
			{
				((IDisposable)_columnIndex[c]).Dispose();
			}
		}

		_values = null;
		_columnIndex = null;
	}

	public bool MoveNext() => GetNextCell(ref _row, ref _colPos, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
	public void Reset()
	{
		_colPos = -1;
		_row = 0;
	}

}








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
 * Jan Källman		    Added       		        2010-02-04
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System.Collections;
using System.Diagnostics.CodeAnalysis;
namespace OfficeOpenXml;

/// <summary>
/// This is the store for all Rows, Columns and Cells.
/// It is a Dictionary implementation that allows you to change the Key (the RowID, ColumnID or CellID )
/// </summary>
internal partial class RangeCollection : IEnumerator<IRangeID>, IEnumerable, IDisposable
{
	private class IndexItem
	{
		internal IndexItem(ulong cellId)
		{
			RangeID = cellId;
		}
		internal IndexItem(ulong cellId, int listPointer)
		{
			RangeID = cellId;
			ListPointer = listPointer;
		}
		internal ulong RangeID;
		internal int ListPointer;
	}
	/// <summary>
	/// Compares an IndexItem
	/// </summary>
	internal class Compare : IComparer<IndexItem>
	{
		int IComparer<IndexItem>.Compare(IndexItem x, IndexItem y) => x.RangeID < y.RangeID ? -1 : x.RangeID > y.RangeID ? 1 : 0;

	}
	IndexItem[] _cellIndex;
	List<IRangeID> _cells;
	static readonly Compare _comparer = new();
	/// <summary>
	/// Creates a new collection
	/// </summary>
	/// <param name="cells">The Cells. This list must be sorted</param>
	internal RangeCollection(List<IRangeID> cells)
	{
		_cells = cells;
		InitSize(_cells);
		for (var i = 0; i < _cells.Count; i++)
		{
			_cellIndex[i] = new IndexItem(cells[i].RangeID, i);
		}
	}
	~RangeCollection()
	{
		_cells = null;
		_cellIndex = null;
	}
	/// <summary>
	/// Return the item with the RangeID
	/// </summary>
	/// <param name="RangeID"></param>
	/// <returns></returns>
	internal IRangeID this[ulong RangeID]
	{
		get
		{
			return _cells[_cellIndex[IndexOf(RangeID)].ListPointer];
		}
	}
	/// <summary>
	/// Return specified index from the sorted list
	/// </summary>
	/// <param name="Index"></param>
	/// <returns></returns>
	internal IRangeID this[int Index]
	{
		get
		{
			return _cells[_cellIndex[Index].ListPointer];
		}
	}
	internal int Count => _cells.Count;
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal void Add(IRangeID cell)
	{
		var ix = IndexOf(cell.RangeID);
		if (ix >= 0)
		{
			throw new Exception("Item already exists");
		}

		Insert(~ix, cell);
	}
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	internal void Delete(ulong key)
	{
		var ix = IndexOf(key);
		if (ix < 0)
		{
			throw new Exception("Key does not exists");
		}

		var listPointer = _cellIndex[ix].ListPointer;
		Array.Copy(_cellIndex, ix + 1, _cellIndex, ix, _cells.Count - ix - 1);
		_cells.RemoveAt(listPointer);

		//Item is removed subtract one from all items with greater ListPointer
		for (var i = 0; i < _cells.Count; i++)
		{
			if (_cellIndex[i].ListPointer >= listPointer)
			{
				_cellIndex[i].ListPointer--;
			}

		}
	}
	internal int IndexOf(ulong key) => Array.BinarySearch<IndexItem>(_cellIndex, 0, _cells.Count, new IndexItem(key), _comparer);
	internal bool ContainsKey(ulong key) => IndexOf(key) >= 0;
	int _size { get; set; }
	/// <summary>
	/// Init the size starting from 128 items. Double the size until the list fits.
	/// </summary>
	/// <param name="_cells"></param>
	private void InitSize(List<IRangeID> _cells)
	{
		_size = 128;
		while (_cells.Count > _size)
		{
			_size <<= 1;
		}

		_cellIndex = new IndexItem[_size];
	}
	/// <summary>
	/// Check the size and double the size if out of bound
	/// </summary>
	private void CheckSize()
	{
		if (_cells.Count >= _size)
		{
			_size <<= 1;
			Array.Resize(ref _cellIndex, _size);
		}
	}
	private void Insert(int ix, IRangeID cell)
	{
		CheckSize();
		Array.Copy(_cellIndex, ix, _cellIndex, ix + 1, _cells.Count - ix);
		_cellIndex[ix] = new IndexItem(cell.RangeID, _cells.Count);
		_cells.Add(cell);
	}

	IRangeID IEnumerator<IRangeID>.Current => throw new NotImplementedException();

	void IDisposable.Dispose() => _ix = -1;

	int _ix = -1;
	object IEnumerator.Current => _cells[_cellIndex[_ix].ListPointer];

	bool IEnumerator.MoveNext()
	{
		_ix++;
		return _ix < _cells.Count;
	}

	void IEnumerator.Reset() => _ix = -1;

	IEnumerator IEnumerable.GetEnumerator() => MemberwiseClone() as IEnumerator;

}

using System.Diagnostics.CodeAnalysis;

[SuppressMessage("Minor Code Smell", "S1210:Equality and comparison operators should be overloaded when implementing IComparable", Justification = "Internal ordering key for Array.BinarySearch; overriding equality would change reference semantics.")]
internal class IndexBase : IComparable<IndexBase>
{
	internal short Index;
	public int CompareTo(IndexBase other) =>
		Index - other.Index;
}
// to compress memory size, use struct
[SuppressMessage("Minor Code Smell", "S1210:Equality and comparison operators should be overloaded when implementing IComparable", Justification = "Internal ordering key for Array.BinarySearch only.")]
[SuppressMessage("Minor Code Smell", "S3898:Value types should implement IEquatable<T>", Justification = "Internal ordering key for Array.BinarySearch only; never compared for equality.")]
internal struct IndexItem : IComparable<IndexItem>
{
	internal int IndexPointer
	{
		get;
		set;
	}
	internal short Index;
	public int CompareTo(IndexItem other) => Index - other.Index;
}
internal class ColumnIndex : IndexBase, IDisposable
{
	internal IndexBase _searchIx = new();
	public ColumnIndex()
	{
		_pages = new PageIndex[CellStore<int>.PagesPerColumnMin];
		PageCount = 0;
	}
	~ColumnIndex()
	{
		_pages = null;
	}
	[SuppressMessage("Code Smell", "S2330:Array covariance should not be used", Justification = "BinarySearch only reads the array, so the covariant view as IndexBase[] cannot cause an ArrayTypeMismatchException.")]
	internal int GetPosition(int Row)
	{
		var page = (short)(Row >> CellStore<int>.pageBits);
		int res;
		if (page >= 0 && page < PageCount && _pages[page].Index == page)
		{
			res = page;
		}
		else
		{
			_searchIx.Index = page;
			res = Array.BinarySearch(_pages, 0, PageCount, _searchIx);
		}

		if (res >= 0)
		{
			GetPage(Row, ref res);
			return res;
		}
		else
		{
			var p = ~res;

			return GetPage(Row, ref p) ? p : res;
		}
	}

	private bool GetPage(int Row, ref int res)
	{
		if (res < PageCount && _pages[res].MinIndex <= Row && _pages[res].MaxIndex >= Row)
		{
			return true;
		}
		else
		{
			if (res + 1 < PageCount && (_pages[res + 1].MinIndex <= Row))
			{
				res = SkipPagesForward(Row, res);
				return true;
			}
			else if (res - 1 >= 0 && _pages[res - 1].MaxIndex >= Row)
			{
				res = SkipPagesBackward(Row, res);
				return true;
			}

			return false;
		}
	}

	private int SkipPagesForward(int Row, int res)
	{
		do
		{
			res++;
		}
		while (res + 1 < PageCount && _pages[res + 1].MinIndex <= Row);
		return res;
	}

	private int SkipPagesBackward(int Row, int res)
	{
		do
		{
			res--;
		}
		while (res - 1 > 0 && _pages[res - 1].MaxIndex >= Row);
		return res;
	}
	internal int GetNextRow(int row)
	{
		var p = GetPosition(row);
		if (p < 0)
		{
			p = ~p;
			if (p >= PageCount)
			{
				return -1;
			}
			else
			{

				if (_pages[p].IndexOffset + _pages[p].Rows[0].Index < row)
				{
					return p + 1 >= PageCount ? -1 : _pages[p + 1].IndexOffset + _pages[p].Rows[0].Index;
				}
				else
				{
					return _pages[p].IndexOffset + _pages[p].Rows[0].Index;
				}
			}
		}
		else
		{
			if (p < PageCount)
			{
				var r = _pages[p].GetNextRow(row);
				if (r >= 0)
				{
					return _pages[p].IndexOffset + _pages[p].Rows[r].Index;
				}
				else
				{
					return ++p < PageCount ? _pages[p].IndexOffset + _pages[p].Rows[0].Index : -1;
				}
			}
			else
			{
				return -1;
			}
		}
	}
	internal int FindNext(int Page)
	{
		var p = GetPosition(Page);
		return p < 0 ? ~p : p;
	}
	internal PageIndex[]? _pages;
	internal int PageCount;

	public void Dispose()
	{
		for (var p = 0; p < PageCount; p++)
		{
			((IDisposable)_pages[p]).Dispose();
		}

		_pages = null;
	}

}

internal class PageIndex : IndexBase, IDisposable
{
	internal IndexItem _searchIx = new();
	public PageIndex()
	{
		Rows = new IndexItem[CellStore<int>.PageSizeMin];
		RowCount = 0;
	}
	public PageIndex(IndexItem[] rows, int count)
	{
		Rows = rows;
		RowCount = count;
	}
	public PageIndex(PageIndex pageItem, int start, int size)
		: this(pageItem, start, size, pageItem.Index, pageItem.Offset)
	{

	}
	public PageIndex(PageIndex pageItem, int start, int size, short index, int offset)
	{
		Rows = new IndexItem[CellStore<int>.GetSize(size)];
		Array.Copy(pageItem.Rows, start, Rows, 0, size);
		RowCount = size;
		Index = index;
		Offset = offset;
	}
	~PageIndex()
	{
		Rows = null;
	}
	internal int Offset;
	internal int IndexOffset => IndexExpanded + Offset;
	internal int IndexExpanded => (Index << CellStore<int>.pageBits);
	internal IndexItem[]? Rows { get; set; }
	internal int RowCount;

	internal int GetPosition(int offset)
	{
		_searchIx.Index = (short)offset;
		return Array.BinarySearch(Rows, 0, RowCount, _searchIx);
	}
	internal int GetNextRow(int row)
	{
		var offset = row - IndexOffset;
		var o = GetPosition(offset);
		if (o < 0)
		{
			o = ~o;
			return o < RowCount ? o : -1;
		}

		return o;
	}

	public int MinIndex => Rows.Length > 0 ? IndexOffset + Rows[0].Index : -1;
	public int MaxIndex => RowCount > 0 ? IndexOffset + Rows[RowCount - 1].Index : -1;
	public int GetIndex(int pos) => IndexOffset + Rows[pos].Index;
	public void Dispose() => Rows = null;
}

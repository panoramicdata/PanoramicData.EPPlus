using OfficeOpenXml;
using System.Collections;

internal class CellsStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
{
	private readonly CellStore<T> _cellStore;
	private int _row, _colPos;
	private readonly int _startRow, _startCol, _endRow, _endCol;
	private int _minRow, _minColPos, _maxRow, _maxColPos;

	public CellsStoreEnumerator(CellStore<T> cellStore) :
		this(cellStore, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
	{
	}

	public CellsStoreEnumerator(CellStore<T> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
	{
		_cellStore = cellStore;

		_startRow = StartRow;
		_startCol = StartCol;
		_endRow = EndRow;
		_endCol = EndCol;

		Init();
	}

	internal void Init()
	{
		_minRow = _startRow;
		_maxRow = _endRow;

		_minColPos = _cellStore.GetPosition(_startCol);
		if (_minColPos < 0)
		{
			_minColPos = ~_minColPos;
		}

		_maxColPos = _cellStore.GetPosition(_endCol);
		if (_maxColPos < 0)
		{
			_maxColPos = ~_maxColPos - 1;
		}

		_row = _minRow;
		_colPos = _minColPos - 1;
	}
	internal int Row => _row;
	internal int Column
	{
		get
		{
			if (_colPos == -1)
			{
				MoveNext();
			}

			return _colPos == -1 ? 0 : _cellStore._columnIndex[_colPos].Index;
		}
	}
	internal T Value
	{
		get
		{
			lock (_cellStore)
			{
				return _cellStore.GetValue(_row, Column);
			}
		}
		set
		{
			lock (_cellStore)
			{
				_cellStore.SetValue(_row, Column, value);
			}
		}
	}
	internal bool Next() =>
		_cellStore.GetNextCell(ref _row, ref _colPos, _minColPos, _maxRow, _maxColPos);
	internal bool Previous()
	{
		lock (_cellStore)
		{
			return _cellStore.GetPrevCell(ref _row, ref _colPos, _minRow, _minColPos, _maxColPos);
		}
	}

	public string CellAddress => ExcelCellBase.GetAddress(Row, Column);

	public IEnumerator<T> GetEnumerator()
	{
		Reset();
		return this;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		Reset();
		return this;
	}

	public T Current => Value;

	public void Dispose()
	{
	}

	object IEnumerator.Current
	{
		get
		{
			Reset();
			return this;
		}
	}

	public bool MoveNext() => Next();

	public void Reset() => Init();
}

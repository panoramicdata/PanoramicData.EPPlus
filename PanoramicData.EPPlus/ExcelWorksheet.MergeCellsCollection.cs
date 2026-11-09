using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Collection containing merged cell addresses
	/// </summary>
	public class MergeCellsCollection : IEnumerable<string>
	{
		internal MergeCellsCollection()
		{
		}

		internal CellStore<int> _cells = new();

		private readonly List<string?> _list = [];

		internal List<string?> List => _list;

		public string this[int row, int column]
		{
			get
			{
				var ix = -1;
				return _cells.Exists(row, column, ref ix) && ix >= 0 && ix < List.Count ? List[ix] : null;
			}
		}
		public string this[int index]
		{
			get
			{
				return _list[index];
			}
		}
		[SuppressMessage("Major Bug", "S2551:Shared resources should not be used for locking", Justification = "Public type: external callers may already synchronise on the collection, so the lock target is kept.")]
		internal void Add(ExcelAddressBase address, bool doValidate)
		{
			int ix;

			//Validate
			if (doValidate && Validate(address) == false)
			{
				throw new ArgumentException("Can't merge and already merged range");
			}

			lock (this)
			{
				ix = _list.Count;
				_list.Add(address.Address);
				SetIndex(address, ix);
			}
		}

		private bool Validate(ExcelAddressBase address)
		{
			var ix = 0;
			if (_cells.Exists(address._fromRow, address._fromCol, ref ix))
			{
				return ix >= 0 && ix < _list.Count && _list[ix] != null && address.Address == _list[ix];
			}

			var cse = new CellsStoreEnumerator<int>(_cells, address._fromRow, address._fromCol, address._toRow, address._toCol);
			//cells
			while (cse.Next())
			{
				return false;
			}
			//Entire column
			cse = new CellsStoreEnumerator<int>(_cells, 0, address._fromCol, 0, address._toCol);
			while (cse.Next())
			{
				return false;
			}
			//Entire row
			cse = new CellsStoreEnumerator<int>(_cells, address._fromRow, 0, address._toRow, 0);
			while (cse.Next())
			{
				return false;
			}

			return true;
		}

		internal void SetIndex(ExcelAddressBase address, int ix)
		{
			int fromRow = address._fromRow, toRow = address._toRow, fromCol = address._fromCol, toCol = address._toCol;
			if (fromRow == 1 && toRow == ExcelPackage.MaxRows) //Entire column, stored on row 0
			{
				fromRow = toRow = 0;
			}
			else if (fromCol == 1 && toCol == ExcelPackage.MaxColumns) //Entire row, stored on column 0
			{
				fromCol = toCol = 0;
			}

			for (var col = fromCol; col <= toCol; col++)
			{
				for (var row = fromRow; row <= toRow; row++)
				{
					_cells.SetValue(row, col, ix);
				}
			}
		}
		public int Count => _list.Count;
		internal void Remove(string Item) => _list.Remove(Item);
		#region IEnumerable<string> Members

		public IEnumerator<string> GetEnumerator() => _list.GetEnumerator();

		#endregion

		#region IEnumerable Members

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => _list.GetEnumerator();

		#endregion
		internal void Clear(ExcelAddressBase Destination)
		{
			var cse = new CellsStoreEnumerator<int>(_cells, Destination._fromRow, Destination._fromCol, Destination._toRow, Destination._toCol);
			var used = new HashSet<int>();
			while (cse.Next())
			{
				var v = cse.Value;
				if (!used.Contains(v) && _list[v] != null)
				{
					var adr = new ExcelAddressBase(_list[v]);
					if (Destination.Collide(adr) is not (ExcelAddressBase.eAddressCollition.Inside or ExcelAddressBase.eAddressCollition.Equal))
					{
						throw new InvalidOperationException(string.Format("Can't delete/overwrite merged cells. A range is partly merged with the another merged range. {0}", adr._address));
					}

					used.Add(v);
				}
			}

			_cells.Clear(Destination._fromRow, Destination._fromCol, Destination._toRow - Destination._fromRow + 1, Destination._toCol - Destination._fromCol + 1);
			foreach (var i in used)
			{
				_list[i] = null;
			}
		}
	}
}

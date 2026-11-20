using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Collection containing merged cell addresses
	/// </summary>
	public partial class MergeCellsCollection : IEnumerable<string>
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

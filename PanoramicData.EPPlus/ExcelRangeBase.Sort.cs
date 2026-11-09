using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Sort the range by value of the first column, Ascending.
	/// </summary>
	public void Sort() => Sort([0], [false]);
	/// <summary>
	/// Sort the range by value of the supplied column, Ascending.
	/// <param name="column">The column to sort by within the range. Zerobased</param>
	/// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
	/// </summary>
	[SuppressMessage("Major Code Smell", "S2360:Optional parameters should not be used", Justification = "Public API: replacing the optional parameter with overloads would break binary compatibility.")]
	public void Sort(int column, bool descending = false) => Sort([column], [descending]);
	/// <summary>
	/// Sort the range by value
	/// </summary>
	/// <param name="columns">The column(s) to sort by within the range. Zerobased</param>
	/// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
	/// <param name="culture">The CultureInfo used to compare values. A null value means CurrentCulture</param>
	/// <param name="compareOptions">String compare option</param>
	[SuppressMessage("Major Code Smell", "S2360:Optional parameters should not be used", Justification = "Public API: replacing the optional parameter with overloads would break binary compatibility.")]
	public void Sort(int[] columns, bool[]? descending = null, CultureInfo? culture = null, CompareOptions compareOptions = CompareOptions.None)
	{
		columns ??= [0];

		var cols = _toCol - _fromCol + 1;
		foreach (var c in columns)
		{
			if (c > cols - 1 || c < 0)
			{
				throw new ArgumentException("Can not reference columns outside the boundries of the range. Note that column reference is zero-based within the range");
			}
		}

		var e = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
		var l = new List<SortItem<ExcelCoreValue>>();
		SortItem<ExcelCoreValue> item = new();

		while (e.Next())
		{
			if (l.Count == 0 || l[^1].Row != e.Row)
			{
				item = new SortItem<ExcelCoreValue> { Row = e.Row, Items = new ExcelCoreValue[cols] };
				l.Add(item);
			}

			item.Items[e.Column - _fromCol] = e.Value;
		}

		if (descending == null)
		{
			descending = new bool[columns.Length];
			for (var i = 0; i < columns.Length; i++)
			{
				descending[i] = false;
			}
		}

		var comp = new Comp
		{
			columns = columns,
			descending = descending,
			cultureInfo = culture ?? CultureInfo.CurrentCulture,
			compareOptions = compareOptions
		};
		l.Sort(comp);

		var flags = GetItems(_worksheet._flags, _fromRow, _fromCol, _toRow, _toCol);
		var formulas = GetItems(_worksheet._formulas, _fromRow, _fromCol, _toRow, _toCol);
		var hyperLinks = GetItems(_worksheet._hyperLinks, _fromRow, _fromCol, _toRow, _toCol);
		var comments = GetItems(_worksheet._commentsStore, _fromRow, _fromCol, _toRow, _toCol);
		//Sort the values and styles.
		_worksheet._values.Clear(_fromRow, _fromCol, _toRow - _fromRow + 1, cols);
		for (var r = 0; r < l.Count; r++)
		{
			for (var c = 0; c < cols; c++)
			{
				var row = _fromRow + r;
				var col = _fromCol + c;
				_worksheet._values.SetValueSpecial(row, col, SortSetValue, l[r].Items[c]);
				var addr = GetAddress(l[r].Row, _fromCol + c);
				//Move flags
				if (flags.TryGetValue(addr, out var value))
				{
					_worksheet._flags.SetValue(row, col, value);
				}
				//Move formulas
				if (formulas.TryGetValue(addr, out var value2))
				{
					_worksheet._formulas.SetValue(row, col, value2);
					if (formulas[addr] is int sfIx)
					{
						var startAddr = new ExcelAddress(Worksheet._sharedFormulas[sfIx].Address);
						var f = Worksheet._sharedFormulas[sfIx];
						if (startAddr._fromRow > row)
						{
							f.Formula = TranslateFromR1C1(TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), row, f.StartCol);
							f.StartRow = row;
							f.Address = GetAddress(row, startAddr._fromCol, startAddr._toRow, startAddr._toCol);
						}
						else if (startAddr._toRow < row)
						{
							f.Address = GetAddress(startAddr._fromRow, startAddr._fromCol, row, startAddr._toCol);
						}
					}
				}

				//Move hyperlinks
				if (hyperLinks.TryGetValue(addr, out var value3))
				{
					_worksheet._hyperLinks.SetValue(row, col, value3);
				}

				//Move comments
				if (comments.TryGetValue(addr, out var i))
				{
					_worksheet._commentsStore.SetValue(row, col, i);
					var comment = _worksheet._comments[i];
					comment.Reference = GetAddress(row, col);
				}
			}
		}
	}

	private static Dictionary<string, T> GetItems<T>(CellStore<T> store, int fromRow, int fromCol, int toRow, int toCol)
	{
		var e = new CellsStoreEnumerator<T>(store, fromRow, fromCol, toRow, toCol);
		var l = new Dictionary<string, T>();
		while (e.Next())
		{
			l.Add(e.CellAddress, e.Value);
		}

		return l;
	}

	private static void SortSetValue(List<ExcelCoreValue> list, int index, object value)
	{
		var v = (ExcelCoreValue)value;
		list[index] = new ExcelCoreValue { _value = v._value, _styleId = v._styleId };
	}
}

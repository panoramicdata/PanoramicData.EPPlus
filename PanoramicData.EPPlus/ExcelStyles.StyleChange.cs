using OfficeOpenXml.Style;
using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	/// <summary>
	/// Handels changes of properties on the style objects
	/// </summary>
	/// <param name="sender"></param>
	/// <param name="e"></param>
	/// <returns></returns>
	internal int PropertyChange(StyleBase sender, Style.StyleChangeEventArgs e)
	{
		var address = new ExcelAddressBase(e.Address);
		var ws = _wb.Worksheets[e.PositionID];
		Dictionary<int, int> styleCashe = [];
		//Set single address
		lock (ws._values)
		{
			SetStyleAddress(sender, e, address, ws, ref styleCashe);
			if (address.Addresses != null)
			{
				//Handle multiaddresses
				foreach (var innerAddress in address.Addresses)
				{
					SetStyleAddress(sender, e, innerAddress, ws, ref styleCashe);
				}
			}
		}

		return 0;
	}
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	private void SetStyleAddress(StyleBase sender, Style.StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, ref Dictionary<int, int> styleCashe)
	{
		if (address.Start.Column == 0 || address.Start.Row == 0)
		{
			throw new Exception("error address");
		}
		//Columns
		else if (address.Start.Row == 1 && address.End.Row == ExcelPackage.MaxRows)
		{
			SetStyleColumns(sender, e, address, ws, styleCashe);
		}

		//Rows
		else if (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns)
		{
			SetStyleRows(sender, e, address, ws, styleCashe);
		}
		else             //Cellrange
		{
			SetStyleCellRange(sender, e, address, ws, styleCashe);
		}
	}

	private int GetOrAddStyleId(StyleBase sender, StyleChangeEventArgs e, Dictionary<int, int> styleCashe, int s)
	{
		if (styleCashe.TryGetValue(s, out var value))
		{
			return value;
		}

		var st = CellXfs[s];
		var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
		styleCashe.Add(s, newId);
		return newId;
	}

	private void SetStyleCellRange(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
	{
		var tmpCache = styleCashe;
		var rowCache = new Dictionary<int, int>(address.End.Row - address.Start.Row + 1);
		var colCache = new Dictionary<int, ExcelCoreValue>(address.End.Column - address.Start.Column + 1);
		ws._values.SetRangeValueSpecial(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column,
			(List<ExcelCoreValue> list, int index, int row, int column, object _) =>
			{
				// Optimized GetStyleID
				var s = list[index]._styleId;
				if (s == 0 && !ws.ExistsStyleInner(row, 0, ref s))
				{
					s = GetCachedRowOrColumnStyleId(ws, rowCache, colCache, row, column);
				}

				list[index] = new ExcelCoreValue { _value = list[index]._value, _styleId = GetOrAddStyleId(sender, e, tmpCache, s) };
			},
			e);

	}

	private static int GetCachedRowOrColumnStyleId(ExcelWorksheet ws, Dictionary<int, int> rowCache, Dictionary<int, ExcelCoreValue> colCache, int row, int column)
	{
		// get row styleId with cache
		if (!rowCache.ContainsKey(row))
		{
			rowCache.Add(row, ws._values.GetValue(row, 0)._styleId);
		}

		var s = rowCache[row];
		if (s == 0)
		{
			// get column styleId with cache
			if (!colCache.ContainsKey(column))
			{
				colCache.Add(column, ws._values.GetValue(0, column));
			}

			s = colCache[column]._styleId;
			if (s == 0)
			{
				int r = 0, c = column;
				if (ws._values.PrevCell(ref r, ref c))
				{
					if (!colCache.ContainsKey(c))
					{
						colCache.Add(c, ws._values.GetValue(0, c));
					}

					var val = colCache[c];
					var colObj = (ExcelColumn)(val._value);
					if (colObj != null && colObj.ColumnMax >= column) //Fixes issue 15174
					{
						s = val._styleId;
					}
				}
			}
		}

		return s;
	}
	internal static int GetStyleId(ExcelWorksheet ws, int row, int col)
	{
		var v = 0;
		if (ws.ExistsStyleInner(row, col, ref v))
		{
			return v;
		}
		else
		{
			if (ws.ExistsStyleInner(row, 0, ref v)) //First Row
			{
				return v;
			}
			else // then column
			{
				if (ws.ExistsStyleInner(0, col, ref v))
				{
					return v;
				}
				else
				{
					int r = 0, c = col;
					if (ws._values.PrevCell(ref r, ref c))
					{
						var val = ws._values.GetValue(0, c);
						var column = (ExcelColumn)(val._value);
						if (column != null && column.ColumnMax >= col) //Fixes issue 15174
						{
							return val._styleId;
						}
						else
						{
							return 0;
						}
					}
					else
					{
						return 0;
					}
				}

			}
		}

	}
	/// <summary>
	/// Handles property changes on Named styles.
	/// </summary>
	/// <param name="sender"></param>
	/// <param name="e"></param>
	/// <returns></returns>
	internal int NamedStylePropertyChange(StyleBase sender, Style.StyleChangeEventArgs e)
	{

		var index = NamedStyles.FindIndexByID(e.Address);
		if (index >= 0)
		{
			var newId = CellStyleXfs[NamedStyles[index].StyleXfId].GetNewID(CellStyleXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
			var prevIx = NamedStyles[index].StyleXfId;
			NamedStyles[index].StyleXfId = newId;
			NamedStyles[index].Style.Index = newId;

			NamedStyles[index].XfId = int.MinValue;
			foreach (var style in CellXfs)
			{
				if (style.XfId == prevIx)
				{
					style.XfId = newId;
				}
			}
		}

		return 0;
	}
}

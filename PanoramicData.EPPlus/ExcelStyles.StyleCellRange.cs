using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
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
}

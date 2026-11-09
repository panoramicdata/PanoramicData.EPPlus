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
			ExcelColumn column;
			int col = address.Start.Column, row = 0;
			bool isNew;
			//Get the startcolumn
			object o = null;
			if (!ws.ExistsValueInner(0, address.Start.Column, ref o))
			{
				column = ws.Column(address.Start.Column);
				isNew = true;
			}
			else
			{
				column = (ExcelColumn)o;
				isNew = false;
			}

			var prevColumMax = column.ColumnMax;
			while (column.ColumnMin <= address.End.Column)
			{
				if (column.ColumnMin > prevColumMax + 1)
				{
					var newColumn = ws.Column(prevColumMax + 1);
					newColumn.ColumnMax = column.ColumnMin - 1;
					AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
				}

				if (column.ColumnMax > address.End.Column)
				{
					ws.CopyColumn(column, address.End.Column + 1, column.ColumnMax);
					column.ColumnMax = address.End.Column;
				}

				var s = ws.GetStyleInner(0, column.ColumnMin);
				AddNewStyleColumn(sender, e, ws, styleCashe, column, s);

				prevColumMax = column.ColumnMax;
				if (!ws._values.NextCell(ref row, ref col) || row > 0)
				{
					if (column._columnMax == address.End.Column)
					{
						break;
					}

					if (isNew)
					{
						column._columnMax = address.End.Column;
					}
					else
					{
						var newColumn = ws.Column(column._columnMax + 1);
						newColumn.ColumnMax = address.End.Column;
						AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
						column = newColumn;
					}

					break;
				}
				else
				{
					column = (ws.GetValueInner(0, col) as ExcelColumn);
				}
			}

			if (column._columnMax < address.End.Column)
			{
				var newCol = ws.Column(column._columnMax + 1) as ExcelColumn;
				newCol._columnMax = address.End.Column;

				var s = ws.GetStyleInner(0, column.ColumnMin);
				if (styleCashe.TryGetValue(s, out var value))
				{
					ws.SetStyleInner(0, column.ColumnMin, value);
				}
				else
				{
					var st = CellXfs[s];
					var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
					styleCashe.Add(s, newId);
					ws.SetStyleInner(0, column.ColumnMin, newId);
				}

				column._columnMax = address.End.Column;
			}

			//Set for individual cells in the span. We loop all cells here since the cells are sorted with columns first.
			var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, address._fromCol, address._toRow, address._toCol);
			while (cse.Next())
			{
				if (cse.Column >= address.Start.Column &&
					cse.Column <= address.End.Column &&
					cse.Value._styleId != 0)
				{
					if (styleCashe.TryGetValue(cse.Value._styleId, out var value))
					{
						ws.SetStyleInner(cse.Row, cse.Column, value);
					}
					else
					{
						var st = CellXfs[cse.Value._styleId];
						var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
						styleCashe.Add(cse.Value._styleId, newId);
						ws.SetStyleInner(cse.Row, cse.Column, newId);
					}
				}
			}

			if (!(address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns))
			{
				//Update cells with styled columns
				cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, 0, address._toRow, 0);
				while (cse.Next())
				{
					if (cse.Value._styleId == 0)
					{
						continue;
					}

					for (var c = address._fromCol; c <= address._toCol; c++)
					{
						if (!ws.ExistsStyleInner(cse.Row, c))
						{
							if (styleCashe.TryGetValue(cse.Value._styleId, out var value))
							{
								ws.SetStyleInner(cse.Row, c, value);
							}
							else
							{
								var st = CellXfs[cse.Value._styleId];
								var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
								styleCashe.Add(cse.Value._styleId, newId);
								ws.SetStyleInner(cse.Row, c, newId);
							}
						}
					}
				}
			}
		}

		//Rows
		else if (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns)
		{
			for (var rowNum = address.Start.Row; rowNum <= address.End.Row; rowNum++)
			{
				var s = ws.GetStyleInner(rowNum, 0);
				if (s == 0)
				{
					//iterate all columns and set the row to the style of the last column
					var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
					while (cse.Next())
					{
						s = cse.Value._styleId;
						if (s == 0)
						{
							continue;
						}

						if (ws.GetValueInner(cse.Row, cse.Column) is ExcelColumn c && c.ColumnMax < ExcelPackage.MaxColumns)
						{
							for (var col = c.ColumnMin; col < c.ColumnMax; col++)
							{
								if (!ws.ExistsStyleInner(rowNum, col))
								{
									ws.SetStyleInner(rowNum, col, s);
								}
							}
						}
					}

					ws.SetStyleInner(rowNum, 0, s);
					cse.Dispose();
				}

				if (styleCashe.TryGetValue(s, out var value))
				{
					ws.SetStyleInner(rowNum, 0, value);
				}
				else
				{
					var st = CellXfs[s];
					var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
					styleCashe.Add(s, newId);
					ws.SetStyleInner(rowNum, 0, newId);
				}
			}

			//Update individual cells 
			var cse2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);
			while (cse2.Next())
			{
				var s = cse2.Value._styleId;
				if (s == 0)
				{
					continue;
				}

				if (styleCashe.TryGetValue(s, out var value))
				{
					ws.SetStyleInner(cse2.Row, cse2.Column, value);
				}
				else
				{
					var st = CellXfs[s];
					var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
					styleCashe.Add(s, newId);
					ws.SetStyleInner(cse2.Row, cse2.Column, newId);
				}
			}

			//Update cells with styled rows
			cse2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, address._toCol);
			while (cse2.Next())
			{
				if (cse2.Value._styleId == 0)
				{
					continue;
				}

				for (var r = address._fromRow; r <= address._toRow; r++)
				{
					if (!ws.ExistsStyleInner(r, cse2.Column))
					{
						var s = cse2.Value._styleId;
						if (styleCashe.TryGetValue(s, out var value))
						{
							ws.SetStyleInner(r, cse2.Column, value);
						}
						else
						{
							var st = CellXfs[s];
							var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
							styleCashe.Add(s, newId);
							ws.SetStyleInner(r, cse2.Column, newId);
						}
					}
				}
			}
		}
		else             //Cellrange
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
						// get row styleId with cache
						if (!rowCache.ContainsKey(row))
						{
							rowCache.Add(row, ws._values.GetValue(row, 0)._styleId);
						}

						s = rowCache[row];
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
					}

					if (tmpCache.TryGetValue(s, out var value))
					{
						list[index] = new ExcelCoreValue { _value = list[index]._value, _styleId = value };
					}
					else
					{
						var st = CellXfs[s];
						var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
						tmpCache.Add(s, newId);
						list[index] = new ExcelCoreValue { _value = list[index]._value, _styleId = newId };
					}
				},
				e);

		}
	}

	private void AddNewStyleColumn(StyleBase sender, StyleChangeEventArgs e, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, int s)
	{
		if (styleCashe.TryGetValue(s, out var value))
		{
			ws.SetStyleInner(0, column.ColumnMin, value);
		}
		else
		{
			var st = CellXfs[s];
			var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
			styleCashe.Add(s, newId);
			ws.SetStyleInner(0, column.ColumnMin, newId);
		}
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

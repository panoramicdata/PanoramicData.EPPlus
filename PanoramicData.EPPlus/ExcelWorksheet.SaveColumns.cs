using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Inserts the cols collection into the XML document
	/// </summary>
	private void UpdateColumnData(StreamWriter sw)
	{
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, ExcelPackage.MaxColumns);
		var first = true;
		while (cse.Next())
		{
			if (first)
			{
				sw.Write("<cols>");
				first = false;
			}

			var col = cse.Value._value as ExcelColumn;
			var cellXfs = _package.Workbook.Styles.CellXfs;

			sw.Write("<col min=\"{0}\" max=\"{1}\"", col.ColumnMin, col.ColumnMax);
			if (col.Hidden == true)
			{
				sw.Write(" hidden=\"1\"");
			}
			else if (col.BestFit)
			{
				sw.Write(" bestFit=\"1\"");
			}

			sw.Write(string.Format(CultureInfo.InvariantCulture, " width=\"{0}\" customWidth=\"1\"", col.Width));

			if (col.OutlineLevel > 0)
			{
				sw.Write(" outlineLevel=\"{0}\" ", col.OutlineLevel);
				if (col.Collapsed)
				{
					if (col.Hidden)
					{
						sw.Write(" collapsed=\"1\"");
					}
					else
					{
						sw.Write(" collapsed=\"1\" hidden=\"1\""); //Always hidden
					}
				}
			}

			if (col.Phonetic)
			{
				sw.Write(" phonetic=\"1\"");
			}

			var styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;
			if (styleID > 0)
			{
				sw.Write(" style=\"{0}\"", styleID);
			}

			sw.Write("/>");
		}

		if (!first)
		{
			sw.Write("</cols>");
		}
	}

	private Dictionary<int, int>? _columnStyles;

	// get StyleID without cell style for UpdateRowCellData
	internal int GetStyleIdDefaultWithMemo(int row, int col)
	{
		var v = 0;
		if (ExistsStyleInner(row, 0, ref v)) //First Row
		{
			return v;
		}
		else // then column
		{
			if (!_columnStyles.ContainsKey(col))
			{
				if (ExistsStyleInner(0, col, ref v))
				{
					_columnStyles.Add(col, v);
				}
				else
				{
					int r = 0, c = col;
					if (_values.PrevCell(ref r, ref c))
					{
						var val = _values.GetValue(0, c);
						var column = (ExcelColumn)(val._value);
						if (column != null && column.ColumnMax >= col) //Fixes issue 15174
						{
							_columnStyles.Add(col, val._styleId);
						}
						else
						{
							_columnStyles.Add(col, 0);
						}
					}
					else
					{
						_columnStyles.Add(col, 0);
					}
				}
			}

			return _columnStyles[col];
		}
	}
}

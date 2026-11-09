using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Utils;
using System.Text;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Insert row and cells into the XML document
	/// </summary>
	private void UpdateRowCellData(StreamWriter sw)
	{
		var cellXfs = _package.Workbook.Styles.CellXfs;

		var row = -1;

		var ss = _package.Workbook._sharedStrings;
		var cache = new StringBuilder();
		cache.Append("<sheetData>");

		FixSharedFormulas(); //Fixes Issue #32

		_columnStyles = [];
		var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
		while (cse.Next())
		{
			if (cse.Column > 0)
			{
				var val = cse.Value;
				var styleID = cellXfs[val._styleId == 0 ? GetStyleIdDefaultWithMemo(cse.Row, cse.Column) : val._styleId].newID;
				//Add the row element if it's a new row
				if (cse.Row != row)
				{
					WriteRow(cache, cellXfs, row, cse.Row);
					row = cse.Row;
				}

				var v = val._value;
				var formula = _formulas.GetValue(cse.Row, cse.Column);
				if (formula is int sfId)
				{
					var f = _sharedFormulas[sfId];
					if (f.Address.IndexOf(':') > 0)
					{
						if (f.StartCol == cse.Column && f.StartRow == cse.Row)
						{
							if (f.IsArray)
							{
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
							}
							else
							{
								cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{6}><f ref=\"{2}\" t=\"shared\" si=\"{3}\">{4}</f>{5}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, sfId, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
							}
						}
						else if (f.IsArray)
						{
							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, styleID < 0 ? 0 : styleID);
						}
						else
						{
							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{4}><f t=\"shared\" si=\"{2}\"/>{3}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, sfId, GetFormulaValue(v), GetCellType(v, true));
						}
					}
					else
					{
						// We can also have a single cell array formula
						if (f.IsArray)
						{
							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, string.Format("{0}:{1}", f.Address, f.Address), ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
						}
						else
						{
							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", f.Address, styleID < 0 ? 0 : styleID, GetCellType(v, true));
							cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v));
						}
					}
				}
				else if (formula != null && formula.ToString() != "")
				{
					cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v, true));
					cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(formula.ToString()), GetFormulaValue(v));
				}
				else
				{
					if (v == null && styleID > 0)
					{
						cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, styleID < 0 ? 0 : styleID);
					}
					else if (v != null)
					{
						// Fix for issue 15460
						if (v is System.Collections.IEnumerable enumerableResult and not string)
						{
							var enumerator = enumerableResult.GetEnumerator();
							if (enumerator.MoveNext() && enumerator.Current != null)
							{
								v = enumerator.Current;
							}
							else
							{
								v = string.Empty;
							}
						}

						if ((TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan))
						{
							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v));
							cache.AppendFormat("{0}</c>", GetFormulaValue(v));
						}
						else
						{
							var vString = Convert.ToString(v);
							int ix;
							if (!ss.TryGetValue(vString, out var value))
							{
								ix = ss.Count;
								ss.Add(vString, new ExcelWorkbook.SharedStringItem { isRichText = _flags.GetFlagValue(cse.Row, cse.Column, CellFlags.RichText), pos = ix });
							}
							else
							{
								ix = value.pos;
							}

							cache.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cse.CellAddress, styleID < 0 ? 0 : styleID);
							cache.AppendFormat("<v>{0}</v></c>", ix);
						}
					}
				}
			}
			else  //ExcelRow
			{
				WriteRow(cache, cellXfs, row, cse.Row);
				row = cse.Row;
			}

			if (cache.Length > 0x600000)
			{
				sw.Write(cache.ToString());
				sw.Flush();
				cache.Length = 0;
			}
		}

		_columnStyles = null;

		if (row != -1)
		{
			cache.Append("</row>");
		}

		cache.Append("</sheetData>");
		sw.Write(cache.ToString());
		sw.Flush();
	}
}

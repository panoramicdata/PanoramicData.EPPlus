using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System.Diagnostics.CodeAnalysis;
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
				WriteCell(cache, cse, cellXfs, ref row, ss);
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

	private void WriteCell(StringBuilder cache, CellsStoreEnumerator<ExcelCoreValue> cse, ExcelStyleCollection<ExcelXfs> cellXfs, ref int row, Dictionary<string, ExcelWorkbook.SharedStringItem> ss)
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
			WriteSharedFormulaCell(cache, cse, styleID, sfId, v);
		}
		else if (formula != null && formula.ToString() != "")
		{
			cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v, true));
			cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(formula.ToString()), GetFormulaValue(v));
		}
		else
		{
			WriteValueCell(cache, cse, styleID, v, ss);
		}
	}

	[SuppressMessage("Minor Bug", "S2692:\"IndexOf\" checks should not be for positive numbers", Justification = "Behaviour-preserving: an address whose first character is ':' is not treated as a range.")]
	private void WriteSharedFormulaCell(StringBuilder cache, CellsStoreEnumerator<ExcelCoreValue> cse, int styleID, int sfId, object v)
	{
		var f = _sharedFormulas[sfId];
		var s = styleID < 0 ? 0 : styleID;
		if (f.Address.IndexOf(':') > 0)
		{
			WriteRangeSharedFormulaCell(cache, cse, s, f, sfId, v);
		}
		else
		{
			// We can also have a single cell array formula
			if (f.IsArray)
			{
				cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, s, string.Format("{0}:{1}", f.Address, f.Address), ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
			}
			else
			{
				cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", f.Address, s, GetCellType(v, true));
				cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v));
			}
		}
	}

	private void WriteRangeSharedFormulaCell(StringBuilder cache, CellsStoreEnumerator<ExcelCoreValue> cse, int s, Formulas f, int sfId, object v)
	{
		if (f.StartCol == cse.Column && f.StartRow == cse.Row)
		{
			if (f.IsArray)
			{
				cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, s, f.Address, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
			}
			else
			{
				cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{6}><f ref=\"{2}\" t=\"shared\" si=\"{3}\">{4}</f>{5}</c>", cse.CellAddress, s, f.Address, sfId, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v, true));
			}
		}
		else if (f.IsArray)
		{
			cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, s);
		}
		else
		{
			cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{4}><f t=\"shared\" si=\"{2}\"/>{3}</c>", cse.CellAddress, s, sfId, GetFormulaValue(v), GetCellType(v, true));
		}
	}
}

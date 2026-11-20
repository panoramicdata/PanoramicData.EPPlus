using OfficeOpenXml.Style.XmlAccess;
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

}

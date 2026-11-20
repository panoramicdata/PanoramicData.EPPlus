using System.Text;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void WriteValueCell(StringBuilder cache, CellsStoreEnumerator<ExcelCoreValue> cse, int styleID, object v, Dictionary<string, ExcelWorkbook.SharedStringItem> ss)
	{
		var s = styleID < 0 ? 0 : styleID;
		if (v == null && styleID > 0)
		{
			cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, s);
		}
		else if (v != null)
		{
			v = GetFirstItemIfEnumerable(v);

			if (IsNumericOrDateValue(v))
			{
				cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, s, GetCellType(v));
				cache.AppendFormat("{0}</c>", GetFormulaValue(v));
			}
			else
			{
				WriteSharedStringCell(cache, cse, s, v, ss);
			}
		}
	}

	private static object GetFirstItemIfEnumerable(object v)
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

		return v;
	}

	private void WriteSharedStringCell(StringBuilder cache, CellsStoreEnumerator<ExcelCoreValue> cse, int s, object v, Dictionary<string, ExcelWorkbook.SharedStringItem> ss)
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

		cache.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cse.CellAddress, s);
		cache.AppendFormat("<v>{0}</v></c>", ix);
	}
}

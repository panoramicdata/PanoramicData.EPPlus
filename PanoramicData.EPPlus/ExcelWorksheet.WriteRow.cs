using OfficeOpenXml.Style.XmlAccess;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void WriteRow(StringBuilder cache, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
	{
		if (prevRow != -1)
		{
			cache.Append("</row>");
		}
		cache.AppendFormat("<row r=\"{0}\"", row);
		if (GetValueInner(row, 0) is RowInternal currRow)
		{
			// if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
			if (currRow.Hidden == true)
			{
				cache.Append(" hidden=\"1\"");
			}

			if (currRow.Height >= 0)
			{
				cache.Append(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));
				if (currRow.CustomHeight)
				{
					cache.Append(" customHeight=\"1\"");
				}
			}

			if (currRow.OutlineLevel > 0)
			{
				cache.AppendFormat(" outlineLevel =\"{0}\"", currRow.OutlineLevel);
				if (currRow.Collapsed)
				{
					if (currRow.Hidden)
					{
						cache.Append(" collapsed=\"1\"");
					}
					else
					{
						cache.Append(" collapsed=\"1\" hidden=\"1\""); //Always hidden
					}
				}
			}

			if (currRow.Phonetic)
			{
				cache.Append(" ph=\"1\"");
			}
		}

		var s = GetStyleInner(row, 0);
		if (s > 0)
		{
			cache.AppendFormat(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID);
		}

		cache.Append('>');
	}
}

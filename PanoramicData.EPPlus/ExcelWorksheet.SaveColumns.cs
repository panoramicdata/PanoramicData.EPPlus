using OfficeOpenXml.Style.XmlAccess;
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

			WriteColumnElement(sw, col, cellXfs);
		}

		if (!first)
		{
			sw.Write("</cols>");
		}
	}

	private static void WriteColumnElement(StreamWriter sw, ExcelColumn col, ExcelStyleCollection<ExcelXfs> cellXfs)
	{
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

		WriteColumnOutline(sw, col);

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

	private static void WriteColumnOutline(StreamWriter sw, ExcelColumn col)
	{
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
	}

}

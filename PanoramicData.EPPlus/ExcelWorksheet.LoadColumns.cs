using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void LoadColumns(XmlReader xr)
	{
		if (ReadUntil(xr, "cols", "sheetData"))
		{
			while (xr.Read())
			{
				if (xr.NodeType == XmlNodeType.Whitespace)
				{
					continue;
				}

				if (xr.LocalName != "col")
				{
					break;
				}

				if (xr.NodeType == XmlNodeType.Element)
				{
					LoadColumn(xr);
				}
			}
		}
	}

	private void LoadColumn(XmlReader xr)
	{
		var min = int.Parse(xr.GetAttribute("min"));

		ExcelColumn col = new(this, min)
		{
			ColumnMax = int.Parse(xr.GetAttribute("max")),
			Width = xr.GetAttribute("width") == null ? 0 : double.Parse(xr.GetAttribute("width"), CultureInfo.InvariantCulture),
			BestFit = xr.GetAttribute("bestFit") is not null and "1",
			Collapsed = xr.GetAttribute("collapsed") is not null and "1",
			Phonetic = xr.GetAttribute("phonetic") is not null and "1",
			OutlineLevel = (short)(xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture)),
			Hidden = xr.GetAttribute("hidden") is not null and "1"
		};
		SetValueInner(0, min, col);

		if (!(xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), NumberStyles.Number, CultureInfo.InvariantCulture, out var style)))
		{
			SetStyleInner(0, min, style);
		}
	}

	/// <summary>
	/// Read until the node is found. If not found the xmlreader is reseted.
	/// </summary>
	/// <param name="xr">The reader</param>
	/// <param name="nodeText">Text to search for</param>
	/// <param name="altNode">Alternative text to search for</param>
	/// <returns></returns>
	private static bool ReadXmlReaderUntil(XmlReader xr, string nodeText, string altNode)
	{
		do
		{
			if (xr.LocalName == nodeText || xr.LocalName == altNode)
			{
				return true;
			}
		}
		while (xr.Read());
		return false;
	}
}

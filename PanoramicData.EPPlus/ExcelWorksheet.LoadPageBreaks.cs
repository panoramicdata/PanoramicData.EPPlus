using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Get the lenth of the attributes
	/// Conditional formatting attributes can be extremly long som get length of the attributes to finetune position.
	/// </summary>
	/// <param name="xr"></param>
	/// <returns></returns>
	private static int GetAttributeLength(XmlReader xr)
	{
		if (xr.NodeType != XmlNodeType.Element)
		{
			return 0;
		}

		var length = 0;

		for (var i = 0; i < xr.AttributeCount; i++)
		{
			var a = xr.GetAttribute(i);
			length += string.IsNullOrEmpty(a) ? 0 : a.Length;
		}

		return length;
	}
	private void LoadRowPageBreakes(XmlReader xr)
	{
		if (ReadUntil(xr, "rowBreaks", "colBreaks"))
		{
			ReadPageBreaks(xr, id => Row(id).PageBreak = true);
		}
	}
	private void LoadColPageBreakes(XmlReader xr)
	{
		if (ReadUntil(xr, "colBreaks"))
		{
			ReadPageBreaks(xr, id => Column(id).PageBreak = true);
		}
	}

	private static void ReadPageBreaks(XmlReader xr, Action<int> setPageBreak)
	{
		while (xr.Read() && xr.LocalName == "brk")
		{
			if (xr.NodeType == XmlNodeType.Element && int.TryParse(xr.GetAttribute("id"), NumberStyles.Number, CultureInfo.InvariantCulture, out var id))
			{
				setPageBreak(id);
			}
		}
	}

	private void ClearNodes()
	{
		_worksheetXml.SelectSingleNode("//d:cols", NameSpaceManager)?.RemoveAll();

		_worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager)?.RemoveAll();

		_worksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager)?.RemoveAll();

		_worksheetXml.SelectSingleNode("//d:rowBreaks", NameSpaceManager)?.RemoveAll();

		_worksheetXml.SelectSingleNode("//d:colBreaks", NameSpaceManager)?.RemoveAll();
	}
}

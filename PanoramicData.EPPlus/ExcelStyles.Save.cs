using OfficeOpenXml.Style.XmlAccess;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	private void AppendUsedStyles<T>(XmlNode parent, ExcelStyleCollection<T> styles, string elementName, int count) where T : StyleXmlHelper
	{
		foreach (var style in styles)
		{
			if (style.useCnt > 0)
			{
				parent.AppendChild(style.CreateXmlNode(_styleXml.CreateElement(elementName, ExcelPackage.schemaMain)));
				style.newID = count;
				count++;
			}
		}

		(parent as XmlElement).SetAttribute("count", count.ToString());
	}

	public void UpdateXml()
	{
		RemoveUnusedStyles();

		var normalIx = NamedStyles.FindIndexByID("Normal");
		UpdateNumberFormatsXml(normalIx);
		UpdateFontsXml(normalIx);

		//Fills
		var fillsNode = _styleXml.SelectSingleNode(FillsPath, _nameSpaceManager);
		fillsNode.RemoveAll();
		Fills[0].useCnt = 1;    // Must exist (none)
		Fills[1].useCnt = 1;    // Must exist (gray125)
		AppendUsedStyles(fillsNode, Fills, "fill", 0);

		//Borders
		var bordersNode = _styleXml.SelectSingleNode(BordersPath, _nameSpaceManager);
		bordersNode.RemoveAll();
		Borders[0].useCnt = 1;    // Must exist blank
		AppendUsedStyles(bordersNode, Borders, "border", 0);

		var count = UpdateNamedStylesXml(normalIx, out var cellXfsNode);
		UpdateCellXfsXml(normalIx, count, cellXfsNode);
		UpdateDxfsXml();
	}

	private void UpdateNumberFormatsXml(int normalIx)
	{
		//NumberFormat
		var nfNode = _styleXml.SelectSingleNode(NumberFormatsPath, _nameSpaceManager);
		if (nfNode == null)
		{
			CreateNode(NumberFormatsPath, true);
			nfNode = _styleXml.SelectSingleNode(NumberFormatsPath, _nameSpaceManager);
		}
		else
		{
			nfNode.RemoveAll();
		}

		var count = 0;
		if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Numberformat.NumFmtID >= 164)
		{
			var nf = NumberFormats[NumberFormats.FindIndexByID(NamedStyles[normalIx].Style.Numberformat.Id)];
			nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
			nf.newID = count++;
		}

		foreach (var nf in NumberFormats)
		{
			if (!nf.BuildIn) //Buildin formats are not updated.
			{
				nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
				nf.newID = count;
				count++;
			}
		}

		(nfNode as XmlElement).SetAttribute("count", count.ToString());
	}

	private void UpdateFontsXml(int normalIx)
	{
		//Font
		var count = 0;
		var fntNode = _styleXml.SelectSingleNode(FontsPath, _nameSpaceManager);
		fntNode.RemoveAll();

		//Normal should be first in the collection
		if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Font.Index > 0)
		{
			var fnt = Fonts[NamedStyles[normalIx].Style.Font.Index];
			fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
			fnt.newID = count++;
		}

		AppendUsedStyles(fntNode, Fonts, "font", count);
	}

	private int UpdateNamedStylesXml(int normalIx, out XmlNode cellXfsNode)
	{
		var styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _nameSpaceManager);
		if (styleXfsNode == null && NamedStyles.Count > 0)
		{
			CreateNode(CellStyleXfsPath);
			styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _nameSpaceManager);
		}

		if (NamedStyles.Count > 0)
		{
			styleXfsNode.RemoveAll();
		}
		//NamedStyles
		var count = normalIx > -1 ? 1 : 0;  //If we have a normal style, we make sure it's added first.

		var cellStyleNode = _styleXml.SelectSingleNode(CellStylesPath, _nameSpaceManager);
		cellStyleNode?.RemoveAll();

		cellXfsNode = _styleXml.SelectSingleNode(CellXfsPath, _nameSpaceManager);
		cellXfsNode.RemoveAll();

		if (NamedStyles.Count > 0 && normalIx >= 0)
		{
			NamedStyles[normalIx].newID = 0;
			AddNamedStyle(0, styleXfsNode, cellXfsNode, NamedStyles[normalIx]);
		}

		foreach (var style in NamedStyles)
		{
			if (!style.Name.Equals("normal", StringComparison.OrdinalIgnoreCase))
			{
				AddNamedStyle(count++, styleXfsNode, cellXfsNode, style);
			}
			else
			{
				style.newID = 0;
			}

			cellStyleNode.AppendChild(style.CreateXmlNode(_styleXml.CreateElement("cellStyle", ExcelPackage.schemaMain)));
		}

		if (cellStyleNode != null)
		{
			(cellStyleNode as XmlElement).SetAttribute("count", count.ToString());
		}

		if (styleXfsNode != null)
		{
			(styleXfsNode as XmlElement).SetAttribute("count", count.ToString());
		}

		return count;
	}
	private void AddNamedStyle(int id, XmlNode styleXfsNode, XmlNode cellXfsNode, ExcelNamedStyleXml style)
	{
		var styleXfs = CellStyleXfs[style.StyleXfId];
		styleXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
		styleXfs.newID = id;
		styleXfs.XfId = style.StyleXfId;

		var ix = CellXfs.FindIndexByID(styleXfs.Id);
		if (ix < 0)
		{
			cellXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
		}
		else
		{
			if (id < 0)
			{
				CellXfs[ix].XfId = id;
			}

			cellXfsNode.AppendChild(CellXfs[ix].CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
			CellXfs[ix].useCnt = 0;
			CellXfs[ix].newID = id;
		}

		if (style.XfId >= 0)
		{
			style.XfId = CellXfs[style.XfId].newID;
		}
		else
		{
			style.XfId = 0;
		}
	}
}

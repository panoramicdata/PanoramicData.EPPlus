using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Style.XmlAccess;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class ExcelStyles
{
	private void UpdateCellXfsXml(int normalIx, int count, XmlNode cellXfsNode)
	{
		//CellStyle
		var xfix = 0;
		foreach (var xf in CellXfs)
		{
			if (xf.useCnt > 0 && !(normalIx >= 0 && NamedStyles[normalIx].StyleXfId == xfix))
			{
				cellXfsNode.AppendChild(xf.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
				xf.newID = count;
				count++;
			}

			xfix++;
		}

		(cellXfsNode as XmlElement).SetAttribute("count", count.ToString());
	}

	private void UpdateDxfsXml()
	{
		//Set dxf styling for conditional Formatting
		var dxfsNode = _styleXml.SelectSingleNode(dxfsPath, _nameSpaceManager);
		foreach (var ws in _wb.Worksheets)
		{
			if (ws is ExcelChartsheet)
			{
				continue;
			}

			foreach (var cf in ws.ConditionalFormatting)
			{
				if (cf.Style.HasValue)
				{
					var ix = Dxfs.FindIndexByID(cf.Style.Id);
					if (ix < 0)
					{
						((ExcelConditionalFormattingRule)cf).DxfId = Dxfs.Count;
						Dxfs.Add(cf.Style.Id, cf.Style);
						var elem = ((XmlDocument)TopNode).CreateElement("d", "dxf", ExcelPackage.schemaMain);
						cf.Style.CreateNodes(new XmlHelperInstance(NameSpaceManager, elem), "");
						dxfsNode.AppendChild(elem);
					}
					else
					{
						((ExcelConditionalFormattingRule)cf).DxfId = ix;
					}
				}
			}
		}

		if (dxfsNode != null)
		{
			(dxfsNode as XmlElement).SetAttribute("count", Dxfs.Count.ToString());
		}
	}

	private void RemoveUnusedStyles()
	{
		CellXfs[0].useCnt = 1; //First item is allways used.
		foreach (var sheet in _wb.Worksheets)
		{
			var cse = new CellsStoreEnumerator<ExcelCoreValue>(sheet._values);
			while (cse.Next())
			{
				var v = cse.Value._styleId;
				if (v >= 0)
				{
					CellXfs[v].useCnt++;
				}
			}
		}

		foreach (var ns in NamedStyles)
		{
			CellStyleXfs[ns.StyleXfId].useCnt++;
		}

		CountUsedStyleParts(CellXfs);
		CountUsedStyleParts(CellStyleXfs);
	}

	private void CountUsedStyleParts(ExcelStyleCollection<ExcelXfs> xfs)
	{
		foreach (var xf in xfs)
		{
			if (xf.useCnt > 0)
			{
				if (xf.FontId >= 0)
				{
					Fonts[xf.FontId].useCnt++;
				}

				if (xf.FillId >= 0)
				{
					Fills[xf.FillId].useCnt++;
				}

				if (xf.BorderId >= 0)
				{
					Borders[xf.BorderId].useCnt++;
				}
			}
		}
	}
}

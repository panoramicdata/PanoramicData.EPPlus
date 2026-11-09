using System.Security;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Update xml with hyperlinks 
	/// </summary>
	/// <param name="sw">The stream</param>
	private void UpdateHyperLinks(StreamWriter sw)
	{
		Dictionary<string, string> hyps = [];
		var cse = new CellsStoreEnumerator<Uri>(_hyperLinks);
		var first = true;
		while (cse.Next())
		{
			var uri = _hyperLinks.GetValue(cse.Row, cse.Column);
			if (first && uri != null)
			{
				sw.Write("<hyperlinks>");
				first = false;
			}

			if (uri is ExcelHyperLink && !string.IsNullOrEmpty((uri as ExcelHyperLink).ReferenceAddress))
			{
				var hl = uri as ExcelHyperLink;
				sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\"{2}{3}/>",
						Cells[cse.Row, cse.Column, cse.Row + hl.RowSpann, cse.Column + hl.ColSpann].Address,
						ExcelCellBase.GetFullAddress(SecurityElement.Escape(Name), SecurityElement.Escape(hl.ReferenceAddress)),
							string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
							string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
			}
			else if (uri != null)
			{
				string id;
				Uri hyp;
				if (uri is ExcelHyperLink link)
				{
					hyp = link.OriginalUri;
				}
				else
				{
					hyp = uri;
				}

				if (hyps.TryGetValue(hyp.OriginalString, out var value))
				{
					id = value;
				}
				else
				{
					var relationship = Part.CreateRelationship(hyp, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
					if (uri is ExcelHyperLink)
					{
						var hl = uri as ExcelHyperLink;
						sw.Write("<hyperlink ref=\"{0}\"{2}{3} r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id,
							string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
							string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
					}
					else
					{
						sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id);
					}

					id = relationship.Id;
				}
			}
		}

		if (!first)
		{
			sw.Write("</hyperlinks>");
		}
	}
	/// <summary>
	/// Create the hyperlinks node in the XML
	/// </summary>
	/// <returns></returns>
	private XmlNode CreateHyperLinkCollection()
	{
		var hl = _worksheetXml.CreateElement("hyperlinks", ExcelPackage.schemaMain);
		var prevNode = _worksheetXml.SelectSingleNode("//d:conditionalFormatting", NameSpaceManager);
		if (prevNode == null)
		{
			prevNode = _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager);
			prevNode ??= _worksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
		}

		return _worksheetXml.DocumentElement.InsertAfter(hl, prevNode);
	}
}

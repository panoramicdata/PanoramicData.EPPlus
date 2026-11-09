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
				WriteInternalHyperLink(sw, cse.Row, cse.Column, uri as ExcelHyperLink);
			}
			else if (uri != null)
			{
				WriteExternalHyperLink(sw, hyps, cse.Row, cse.Column, uri);
			}
		}

		if (!first)
		{
			sw.Write("</hyperlinks>");
		}
	}

	private void WriteInternalHyperLink(StreamWriter sw, int row, int column, ExcelHyperLink hl) =>
		sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\"{2}{3}/>",
				Cells[row, column, row + hl.RowSpann, column + hl.ColSpann].Address,
				ExcelCellBase.GetFullAddress(SecurityElement.Escape(Name), SecurityElement.Escape(hl.ReferenceAddress)),
					string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
					string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");

	private void WriteExternalHyperLink(StreamWriter sw, Dictionary<string, string> hyps, int row, int column, Uri uri)
	{
		Uri hyp;
		if (uri is ExcelHyperLink link)
		{
			hyp = link.OriginalUri;
		}
		else
		{
			hyp = uri;
		}

		if (!hyps.ContainsKey(hyp.OriginalString))
		{
			var relationship = Part.CreateRelationship(hyp, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
			if (uri is ExcelHyperLink hl)
			{
				sw.Write("<hyperlink ref=\"{0}\"{2}{3} r:id=\"{1}\"/>", ExcelCellBase.GetAddress(row, column), relationship.Id,
					string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
					string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
			}
			else
			{
				sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\"/>", ExcelCellBase.GetAddress(row, column), relationship.Id);
			}
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

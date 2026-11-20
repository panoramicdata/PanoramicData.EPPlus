using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Load Hyperlinks
	/// </summary>
	/// <param name="xr">The reader</param>
	private void LoadHyperLinks(XmlReader xr)
	{
		if (!ReadUntil(xr, "hyperlinks", "rowBreaks", "colBreaks"))
		{
			return;
		}

		while (xr.Read())
		{
			if (xr.LocalName == "hyperlink")
			{
				ExcelCellBase.GetRowColFromAddress(xr.GetAttribute("ref"), out var fromRow, out var fromCol, out var toRow, out int toCol);
				ExcelHyperLink hl = null;
				if (xr.GetAttribute("id", ExcelPackage.schemaRelationships) != null)
				{
					hl = LoadHyperLinkFromRelationship(xr);
				}
				else if (xr.GetAttribute("location") != null)
				{
					hl = GetHyperlinkFromRef(xr, "location", fromRow, toRow, fromCol, toCol);
				}
				else if (xr.GetAttribute("ref") != null)
				{
					hl = GetHyperlinkFromRef(xr, "ref", fromRow, toRow, fromCol, toCol);
				}
				else
				{
					// not enough info to create a hyperlink
					break;
				}

				var tt = xr.GetAttribute("tooltip");
				if (!string.IsNullOrEmpty(tt))
				{
					hl.ToolTip = tt;
				}

				_hyperLinks.SetValue(fromRow, fromCol, hl);
			}
			else
			{
				break;
			}
		}
	}

	private ExcelHyperLink LoadHyperLinkFromRelationship(XmlReader xr)
	{
		ExcelHyperLink hl;
		var rId = xr.GetAttribute("id", ExcelPackage.schemaRelationships);
		var uri = Part.GetRelationship(rId).TargetUri;
		if (uri.IsAbsoluteUri)
		{
			try
			{
				hl = new ExcelHyperLink(uri.AbsoluteUri);
			}
			catch
			{
				hl = new ExcelHyperLink(uri.OriginalString, UriKind.Absolute);
			}
		}
		else
		{
			hl = new ExcelHyperLink(uri.OriginalString, UriKind.Relative);
		}

		hl.RId = rId;
		Part.DeleteRelationship(rId); //Delete the relationship, it is recreated when we save the package.
		return hl;
	}

	private static ExcelHyperLink GetHyperlinkFromRef(XmlReader xr, string refTag, int fromRow = 0, int toRow = 0, int fromCol = 0, int toCol = 0)
	{
		var hl = new ExcelHyperLink(xr.GetAttribute(refTag), xr.GetAttribute("display"))
		{
			RowSpann = toRow - fromRow,
			ColSpann = toCol - fromCol
		};
		return hl;
	}
}

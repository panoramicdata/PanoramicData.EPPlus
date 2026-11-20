using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Lazily creates the fill, border and font wrappers shared by chart elements.
/// </summary>
internal static class ChartStyleAccess
{
	internal static ExcelDrawingFill GetFill(XmlHelper owner, ref ExcelDrawingFill fill) =>
		fill ??= new ExcelDrawingFill(owner.NameSpaceManager, owner.TopNode, "c:spPr");

	internal static ExcelDrawingBorder GetBorder(XmlHelper owner, ref ExcelDrawingBorder border) =>
		border ??= new ExcelDrawingBorder(owner.NameSpaceManager, owner.TopNode, "c:spPr/a:ln");

	internal static ExcelTextFont GetFont(XmlHelper owner, ref ExcelTextFont font, string[] schemaNodeOrder)
	{
		if (font == null)
		{
			if (owner.TopNode.SelectSingleNode("c:txPr", owner.NameSpaceManager) == null)
			{
				owner.CreateNode("c:txPr/a:bodyPr");
				owner.CreateNode("c:txPr/a:lstStyle");
			}

			font = new ExcelTextFont(owner.NameSpaceManager, owner.TopNode, "c:txPr/a:p/a:pPr/a:defRPr", schemaNodeOrder);
		}

		return font;
	}
}

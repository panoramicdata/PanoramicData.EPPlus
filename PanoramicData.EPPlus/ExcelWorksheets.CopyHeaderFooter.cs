using OfficeOpenXml.Drawing.Vml;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	private void CopyHeaderFooterPictures(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		if (Copy.TopNode != null && Copy.TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager) == null)
		{
			return;
		}
		CopyHeaderFooterTexts(Copy, added);

		// Copy any images
		if (Copy.HeaderFooter.Pictures.Count > 0)
		{
			added.DeleteNode("d:legacyDrawingHF");

			foreach (ExcelVmlDrawingPicture pic in Copy.HeaderFooter.Pictures)
			{
				var item = added.HeaderFooter.Pictures.Add(pic.Id, pic.ImageUri, pic.Title, pic.Width, pic.Height);
				foreach (XmlAttribute att in pic.TopNode.Attributes)
				{
					(item.TopNode as XmlElement).SetAttribute(att.Name, att.Value);
				}

				item.TopNode.InnerXml = pic.TopNode.InnerXml;
			}
		}
	}

	private static void CopyHeaderFooterTexts(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		//Copy the texts
		if (Copy.HeaderFooter._oddHeader != null)
		{
			CopyText(Copy.HeaderFooter._oddHeader, added.HeaderFooter.OddHeader);
		}

		if (Copy.HeaderFooter._oddFooter != null)
		{
			CopyText(Copy.HeaderFooter._oddFooter, added.HeaderFooter.OddFooter);
		}

		if (Copy.HeaderFooter._evenHeader != null)
		{
			CopyText(Copy.HeaderFooter._evenHeader, added.HeaderFooter.EvenHeader);
		}

		if (Copy.HeaderFooter._evenFooter != null)
		{
			CopyText(Copy.HeaderFooter._evenFooter, added.HeaderFooter.EvenFooter);
		}

		if (Copy.HeaderFooter._firstHeader != null)
		{
			CopyText(Copy.HeaderFooter._firstHeader, added.HeaderFooter.FirstHeader);
		}

		if (Copy.HeaderFooter._firstFooter != null)
		{
			CopyText(Copy.HeaderFooter._firstFooter, added.HeaderFooter.FirstFooter);
		}
	}

	private static void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
	{
		to.LeftAlignedText = from.LeftAlignedText;
		to.CenteredText = from.CenteredText;
		to.RightAlignedText = from.RightAlignedText;
	}
}

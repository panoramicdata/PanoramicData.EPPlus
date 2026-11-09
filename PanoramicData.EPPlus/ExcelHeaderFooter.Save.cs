using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public sealed partial class ExcelHeaderFooter
{
	/// <summary>
	/// Saves the header and footer information to the worksheet XML
	/// </summary>
	internal void Save()
	{
		SaveText("d:oddHeader", _oddHeader);
		SaveText("d:oddFooter", _oddFooter);

		// only set evenHeader and evenFooter
		if (differentOddEven)
		{
			SaveText("d:evenHeader", _evenHeader);
			SaveText("d:evenFooter", _evenFooter);
		}

		// only set firstHeader and firstFooter
		if (differentFirst)
		{
			SaveText("d:firstHeader", _firstHeader);
			SaveText("d:firstFooter", _firstFooter);
		}
	}

	private void SaveText(string path, ExcelHeaderFooterText? text)
	{
		if (text != null)
		{
			SetXmlNodeString(path, GetText(text));
		}
	}
	internal void SaveHeaderFooterImages()
	{
		if (_vmlDrawingsHF != null)
		{
			if (_vmlDrawingsHF.Count == 0)
			{
				if (_vmlDrawingsHF.Uri != null)
				{
					_ws.Part.DeleteRelationship(_vmlDrawingsHF.RelId);
					_ws._package.Package.DeletePart(_vmlDrawingsHF.Uri);
				}
			}
			else
			{
				if (_vmlDrawingsHF.Uri == null)
				{
					_vmlDrawingsHF.Uri = GetNewUri(_ws._package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
				}

				if (_vmlDrawingsHF.Part == null)
				{
					_vmlDrawingsHF.Part = _ws._package.Package.CreatePart(_vmlDrawingsHF.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _ws._package.Compression);
					var rel = _ws.Part.CreateRelationship(UriHelper.GetRelativeUri(_ws.WorksheetUri, _vmlDrawingsHF.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
					_ws.SetHFLegacyDrawingRel(rel.Id);
					_vmlDrawingsHF.RelId = rel.Id;
					foreach (ExcelVmlDrawingPicture draw in _vmlDrawingsHF)
					{
						rel = _vmlDrawingsHF.Part.CreateRelationship(UriHelper.GetRelativeUri(_vmlDrawingsHF.Uri, draw.ImageUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
						draw.RelId = rel.Id;
					}
				}

				_vmlDrawingsHF.VmlDrawingXml.Save(_vmlDrawingsHF.Part.GetStream());
			}
		}
	}
	private static string GetText(ExcelHeaderFooterText headerFooter)
	{
		var ret = "";
		if (headerFooter.LeftAlignedText != null)
		{
			ret += "&L" + headerFooter.LeftAlignedText;
		}

		if (headerFooter.CenteredText != null)
		{
			ret += "&C" + headerFooter.CenteredText;
		}

		if (headerFooter.RightAlignedText != null)
		{
			ret += "&R" + headerFooter.RightAlignedText;
		}

		return ret;
	}
}

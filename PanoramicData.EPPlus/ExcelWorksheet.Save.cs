using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Packaging.DotNetZip;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	internal void Save()
	{
		DeletePrinterSettings();

		if (_worksheetXml != null)
		{

			if (this is not ExcelChartsheet)
			{
				// save the header & footer (if defined)
				if (_headerFooter != null)
				{
					HeaderFooter.Save();
				}

				var d = Dimension;
				if (d == null)
				{
					DeleteAllNode("d:dimension/@ref");
				}
				else
				{
					SetXmlNodeString("d:dimension/@ref", d.Address);
				}


				if (Drawings.Count == 0)
				{
					//Remove node if no drawings exists.
					DeleteNode("d:drawing");
				}

				SaveComments();
				HeaderFooter.SaveHeaderFooterImages();
				SaveTables();
				SavePivotTables();
			}
		}

		if (Drawings.UriDrawing != null)
		{
			SaveDrawingsPart();
		}
	}

	private void SaveDrawingsPart()
	{
		if (Drawings.Count == 0)
		{
			Part.DeleteRelationship(Drawings._drawingRelation.Id);
			_package.Package.DeletePart(Drawings.UriDrawing);
		}
		else
		{
			foreach (ExcelDrawing d in Drawings)
			{
				d.AdjustPositionAndSize();
				if (d is ExcelChart c)
				{
					c.ChartXml.Save(c.Part.GetStream(FileMode.Create, FileAccess.Write));
				}
			}

			var partPack = Drawings.Part;
			Drawings.DrawingXml.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
		}
	}
	internal void SaveHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
	{
		//Init Zip
		stream.CodecBufferSize = 8096;
		stream.CompressionLevel = (Packaging.DotNetZip.Zlib.CompressionLevel)compressionLevel;
		stream.PutNextEntry(fileName);


		SaveXml(stream);
	}



	/// <summary>
	/// Delete the printersettings relationship and part.
	/// </summary>
	private void DeletePrinterSettings()
	{
		//Delete the relationship from the pageSetup tag
		var attr = (XmlAttribute)WorksheetXml.SelectSingleNode("//d:pageSetup/@r:id", NameSpaceManager);
		if (attr != null)
		{
			var relID = attr.Value;
			//First delete the attribute from the XML
			attr.OwnerElement.Attributes.Remove(attr);
			if (Part.RelationshipExists(relID))
			{
				var rel = Part.GetRelationship(relID);
				var printerSettingsUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
				Part.DeleteRelationship(rel.Id);

				//Delete the part from the package
				if (_package.Package.PartExists(printerSettingsUri))
				{
					_package.Package.DeletePart(printerSettingsUri);
				}
			}
		}
	}
}

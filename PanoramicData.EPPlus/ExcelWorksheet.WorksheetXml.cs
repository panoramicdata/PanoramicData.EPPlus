using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// The XML document holding the worksheet data.
	/// All column, row, cell, pagebreak, merged cell and hyperlink-data are loaded into memory and removed from the document when loading the document.        
	/// </summary>
	public XmlDocument WorksheetXml => (_worksheetXml);

	internal ExcelVmlDrawingCommentCollection? _vmlDrawings;

	/// <summary>
	/// Vml drawings. underlaying object for comments
	/// </summary>
	internal ExcelVmlDrawingCommentCollection VmlDrawingsComments
	{
		get
		{
			if (_vmlDrawings == null)
			{
				CreateVmlCollection();
			}

			return _vmlDrawings;
		}
	}

	internal ExcelCommentCollection? _comments;
	/// <summary>
	/// Collection of comments
	/// </summary>
	public ExcelCommentCollection Comments
	{
		get
		{
			CheckSheetType();
			if (_comments == null)
			{
				CreateVmlCollection();
				_comments = new ExcelCommentCollection(_package, this, NameSpaceManager);
			}

			return _comments;
		}
	}
	private void CreateVmlCollection()
	{
		var vmlNode = _worksheetXml.DocumentElement.SelectSingleNode("d:legacyDrawing/@r:id", NameSpaceManager);
		if (vmlNode == null)
		{
			_vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, null);
		}
		else
		{
			if (Part.RelationshipExists(vmlNode.Value))
			{
				var rel = Part.GetRelationship(vmlNode.Value);
				var vmlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

				_vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, vmlUri)
				{
					RelId = rel.Id
				};
			}
		}
	}

	private void CreateXml()
	{
		_worksheetXml = new XmlDocument
		{
			PreserveWhitespace = ExcelPackage.preserveWhitespace
		};
		var packPart = _package.Package.GetPart(WorksheetUri);
		var xml = "";

		// First Columns, rows, cells, mergecells, hyperlinks and pagebreakes are loaded from a xmlstream to optimize speed...
		var doAdjust = _package.DoAdjustDrawings;
		_package.DoAdjustDrawings = false;
		Stream stream = packPart.GetStream();

		var xr = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = true });
		LoadColumns(xr);    //columnXml
		var start = stream.Position;
		LoadCells(xr);
		var nextElementLength = GetAttributeLength(xr);
		var end = stream.Position - nextElementLength;
		LoadMergeCells(xr);
		LoadHyperLinks(xr);
		LoadRowPageBreakes(xr);
		LoadColPageBreakes(xr);
		//...then the rest of the Xml is extracted and loaded into the WorksheetXml document.
		stream.Seek(0, SeekOrigin.Begin);
		xml = GetWorkSheetXml(stream, start, end, out var encoding);

		// now release stream buffer (already converted whole Xml into XmlDocument Object and String)
		stream.Dispose();
		packPart.Stream = new MemoryStream();

		//first char is invalid sometimes?? 
		if (xml[0] != '<')
		{
			LoadXmlSafe(_worksheetXml, xml[1..], encoding);
		}
		else
		{
			LoadXmlSafe(_worksheetXml, xml, encoding);
		}

		_package.DoAdjustDrawings = doAdjust;
		ClearNodes();
	}
}

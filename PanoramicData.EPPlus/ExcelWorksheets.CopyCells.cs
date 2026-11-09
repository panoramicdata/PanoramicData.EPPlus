using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	private void CloneCells(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		var sameWorkbook = (Copy.Workbook == _pck.Workbook);

		var doAdjust = _pck.DoAdjustDrawings;
		_pck.DoAdjustDrawings = false;
		//Merged cells
		foreach (var r in Copy.MergedCells)     //Issue #94
		{
			added.MergedCells.Add(new ExcelAddress(r), false);
		}

		//Shared Formulas   
		foreach (var key in Copy._sharedFormulas.Keys)
		{
			added._sharedFormulas.Add(key, Copy._sharedFormulas[key].Clone());
		}

		Dictionary<int, int> styleCashe = [];
		//Cells
		int row, col;
		var val = new CellsStoreEnumerator<ExcelCoreValue>(Copy._values);
		while (val.Next())
		{
			row = val.Row;
			col = val.Column;
			var styleID = CloneCell(Copy, added, row, col);

			if (!sameWorkbook)
			{
				if (styleCashe.TryGetValue(styleID, out var value))
				{
					added.SetStyleInner(row, col, value);
				}
				else
				{
					var s = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles, styleID);
					styleCashe.Add(styleID, s);
					added.SetStyleInner(row, col, s);
				}
			}
		}

		added._package.DoAdjustDrawings = doAdjust;
	}

	private static int CloneCell(ExcelWorksheet Copy, ExcelWorksheet added, int row, int col)
	{
		var styleID = 0;
		if (row == 0) //Column
		{
			if (Copy.GetValueInner(row, col) is ExcelColumn c)
			{
				var clone = c.Clone(added, c.ColumnMin);
				clone.StyleID = c.StyleID;
				added.SetValueInner(row, col, clone);
				styleID = c.StyleID;
			}
		}
		else if (col == 0) //Row
		{
			var r = Copy.Row(row);
			if (r != null)
			{
				r.Clone(added);
				styleID = r.StyleID;
			}

		}
		else
		{
			styleID = CopyValues(Copy, added, row, col);
		}

		return styleID;
	}

	private static int CopyValues(ExcelWorksheet Copy, ExcelWorksheet added, int row, int col)
	{
		added.SetValueInner(row, col, Copy.GetValueInner(row, col));
		byte fl = 0;
		if (Copy._flags.Exists(row, col, ref fl))
		{
			added._flags.SetValue(row, col, fl);
		}

		var v = Copy._formulas.GetValue(row, col);
		if (v != null)
		{
			added.SetFormula(row, col, v);
		}

		var s = Copy.GetStyleInner(row, col);
		if (s != 0)
		{
			added.SetStyleInner(row, col, s);
		}

		var f = Copy._formulas.GetValue(row, col);
		if (f != null)
		{
			added._formulas.SetValue(row, col, f);
		}

		return s;
	}

	private void CopyComment(ExcelWorksheet Copy, ExcelWorksheet workSheet)
	{
		//First copy the drawing XML
		var xml = Copy.Comments.CommentXml.InnerXml;
		var uriComment = new Uri(string.Format("/xl/comments{0}.xml", workSheet.SheetID), UriKind.Relative);
		if (_pck.Package.PartExists(uriComment))
		{
			uriComment = GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
		}

		var part = _pck.Package.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _pck.Compression);

		StreamWriter streamDrawing = new(part.GetStream(FileMode.Create, FileAccess.Write));
		streamDrawing.Write(xml);
		streamDrawing.Flush();

		//Add the relationship ID to the worksheet xml.
		workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriComment), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

		xml = Copy.VmlDrawingsComments.VmlDrawingXml.InnerXml;

		var uriVml = new Uri(string.Format("/xl/drawings/vmldrawing{0}.vml", workSheet.SheetID), UriKind.Relative);
		if (_pck.Package.PartExists(uriVml))
		{
			uriVml = GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
		}

		var vmlPart = _pck.Package.CreatePart(uriVml, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
		StreamWriter streamVml = new(vmlPart.GetStream(FileMode.Create, FileAccess.Write));
		streamVml.Write(xml);
		streamVml.Flush();

		var newVmlRel = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriVml), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

		//Add the relationship ID to the worksheet xml.
		var e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
		if (e == null)
		{
			workSheet.CreateNode("d:legacyDrawing");
			e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
		}

		e.SetAttribute("id", ExcelPackage.schemaRelationships, newVmlRel.Id);
	}
	private void CopyDrawing(ExcelWorksheet Copy, ExcelWorksheet workSheet)
	{
		//First copy the drawing XML                
		var xml = Copy.Drawings.DrawingXml.OuterXml;
		var uriDraw = new Uri(string.Format("/xl/drawings/drawing{0}.xml", workSheet.SheetID), UriKind.Relative);
		var part = _pck.Package.CreatePart(uriDraw, "application/vnd.openxmlformats-officedocument.drawing+xml", _pck.Compression);
		StreamWriter streamDrawing = new(part.GetStream(FileMode.Create, FileAccess.Write));
		streamDrawing.Write(xml);
		streamDrawing.Flush();

		XmlDocument drawXml = new();
		drawXml.LoadXml(xml);
		//Add the relationship ID to the worksheet xml.
		var drawRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriDraw), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
		var e = workSheet.WorksheetXml.SelectSingleNode("//d:drawing", _namespaceManager) as XmlElement;
		e.SetAttribute("id", ExcelPackage.schemaRelationships, drawRelation.Id);

		for (var i = 0; i < Copy.Drawings.Count; i++)
		{
			var draw = Copy.Drawings[i];
			draw.AdjustPositionAndSize();       //Adjust position for any change in normal style font/row size etc.
			if (draw is ExcelChart chart)
			{
				CopyChartPart(Copy, part, uriDraw, drawXml, chart);
			}
			else if (draw is ExcelPicture pic)
			{
				CopyPicturePart(Copy, workSheet, part, drawXml, pic);
			}
		}
		//rewrite the drawing xml with the new relID's
		streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
		streamDrawing.Write(drawXml.OuterXml);
		streamDrawing.Flush();

		//Copy the size variables to the copy.
		for (var i = 0; i < Copy.Drawings.Count; i++)
		{
			var draw = Copy.Drawings[i];
			var c = workSheet.Drawings[i];
			if (c != null)
			{
				c._left = draw._left;
				c._top = draw._top;
				c._height = draw._height;
				c._width = draw._width;
			}
		}
	}

	private void CopyChartPart(ExcelWorksheet Copy, Packaging.ZipPackagePart part, Uri uriDraw, XmlDocument drawXml, ExcelChart chart)
	{
		var xml = chart.ChartXml.InnerXml;

		var UriChart = GetNewUri(_pck.Package, "/xl/charts/chart{0}.xml");
		var chartPart = _pck.Package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _pck.Compression);
		StreamWriter streamChart = new(chartPart.GetStream(FileMode.Create, FileAccess.Write));
		streamChart.Write(xml);
		streamChart.Flush();
		//Now create the new relationship to the copied chart xml
		var prevRelID = chart.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", Copy.Drawings.NameSpaceManager).Value;
		var rel = part.CreateRelationship(UriHelper.GetRelativeUri(uriDraw, UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
		var relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), Copy.Drawings.NameSpaceManager) as XmlAttribute;
		relAtt.Value = rel.Id;
	}

	private void CopyPicturePart(ExcelWorksheet Copy, ExcelWorksheet workSheet, Packaging.ZipPackagePart part, XmlDocument drawXml, ExcelPicture pic)
	{
		var uri = pic.UriPic;
		if (!workSheet.Workbook._package.Package.PartExists(uri))
		{
			var picPart = workSheet.Workbook._package.Package.CreatePart(uri, pic.ContentType, CompressionLevel.None);
			var imageBytes = pic.ImageBytes;
			picPart.GetStream(FileMode.Create, FileAccess.Write).Write(imageBytes, 0, imageBytes.Length);
		}

		var rel = part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
		//Fixes problem with invalid image when the same image is used more than once.
		var relAtt =
			drawXml.SelectSingleNode(
				string.Format(
					"//xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name[.='{0}']/../../../xdr:blipFill/a:blip/@r:embed",
					pic.Name), Copy.Drawings.NameSpaceManager);
		if (relAtt != null)
		{
			relAtt.Value = rel.Id;
		}

		if (_pck._images.TryGetValue(pic.ImageHash, out var value))
		{
			value.RefCount++;
		}
	}

	private void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
	{
		var xml = origSheet.VmlDrawingsComments.VmlDrawingXml.OuterXml;
		var vmlUri = new Uri(string.Format("/xl/drawings/vmlDrawing{0}.vml", newSheet.SheetID), UriKind.Relative);
		var part = _pck.Package.CreatePart(vmlUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
		using (var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write)))
		{
			streamDrawing.Write(xml);
			streamDrawing.Flush();
		}

		//Add the relationship ID to the worksheet xml.
		var vmlRelation = newSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newSheet.WorksheetUri, vmlUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
		var e = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
		e ??= newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", _namespaceManager.LookupNamespace("d")) as XmlElement;

		e?.SetAttribute("id", ExcelPackage.schemaRelationships, vmlRelation.Id);
	}
}

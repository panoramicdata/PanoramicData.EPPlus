using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void SaveComments()
	{
		if (_comments != null)
		{
			SaveCommentsPart();
		}

		if (_vmlDrawings != null)
		{
			SaveVmlDrawingsPart();
		}
	}

	private void SaveCommentsPart()
	{
		if (_comments.Count == 0)
		{
			if (_comments.Uri != null)
			{
				Part.DeleteRelationship(_comments.RelId);
				_package.Package.DeletePart(_comments.Uri);
			}

			RemoveLegacyDrawingRel(VmlDrawingsComments.RelId);
		}
		else
		{
			if (_comments.Uri == null)
			{
				var id = SheetID;
				_comments.Uri = GetNewUri(_package.Package, @"/xl/comments{0}.xml", ref id); //Issue 236-Part already exists fix
			}

			if (_comments.Part == null)
			{
				_comments.Part = _package.Package.CreatePart(_comments.Uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _package.Compression);
				Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _comments.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");
			}

			_comments.CommentXml.Save(_comments.Part.GetStream(FileMode.Create));
		}
	}

	private void SaveVmlDrawingsPart()
	{
		if (_vmlDrawings.Count == 0)
		{
			if (_vmlDrawings.Uri != null)
			{
				Part.DeleteRelationship(_vmlDrawings.RelId);
				_package.Package.DeletePart(_vmlDrawings.Uri);
			}
		}
		else
		{
			if (_vmlDrawings.Uri == null)
			{
				var id = SheetID;
				_vmlDrawings.Uri = GetNewUri(_package.Package, @"/xl/drawings/vmlDrawing{0}.vml", ref id);
			}

			if (_vmlDrawings.Part == null)
			{
				_vmlDrawings.Part = _package.Package.CreatePart(_vmlDrawings.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _package.Compression);
				var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _vmlDrawings.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
				SetXmlNodeString("d:legacyDrawing/@r:id", rel.Id);
				_vmlDrawings.RelId = rel.Id;
			}

			_vmlDrawings.VmlDrawingXml.Save(_vmlDrawings.Part.GetStream(FileMode.Create));
		}
	}
}

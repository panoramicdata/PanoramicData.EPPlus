using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	/// <summary>
	/// Deletes a worksheet from the collection
	/// </summary>
	/// <param name="Index">The position of the worksheet in the workbook</param>
	public void Delete(int Index)
	{
		/*
            * Hack to prefetch all the drawings,
            * so that all the images are referenced, 
            * to prevent the deletion of the image file, 
            * when referenced more than once
            */
		foreach (var ws in _worksheets)
		{
			_ = ws.Value.Drawings;
		}

		var worksheet = _worksheets[Index];
		if (worksheet.Drawings.Count > 0)
		{
			worksheet.Drawings.ClearDrawings();
		}

		//Remove all comments
		if (worksheet is not ExcelChartsheet && worksheet.Comments.Count > 0)
		{
			worksheet.Comments.Clear();
		}

		//Delete any parts still with relations to the Worksheet.
		DeleteRelationsAndParts(worksheet.Part);

		//Delete the worksheet part and relation from the package 
		_pck.Workbook.Part.DeleteRelationship(worksheet.RelationshipID);

		//Delete worksheet from the workbook XML
		var sheetsNode = _pck.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", _namespaceManager);
		if (sheetsNode != null)
		{
			var sheetNode = sheetsNode.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetID), _namespaceManager);
			if (sheetNode != null)
			{
				sheetsNode.RemoveChild(sheetNode);
			}
		}

		_worksheets.Remove(Index);
		_pck.Workbook.VbaProject?.Modules.Remove(worksheet.CodeModule);

		ReindexWorksheetDictionary();
		UpdateActiveTabAfterDelete(worksheet.SheetID);
	}

	private void UpdateActiveTabAfterDelete(int deletedSheetId)
	{
		//If the active sheet is deleted, set the first tab as active.
		if (_pck.Workbook.View.ActiveTab >= _pck.Workbook.Worksheets.Count)
		{
			_pck.Workbook.View.ActiveTab = _pck.Workbook.View.ActiveTab - 1;
		}

		if (_pck.Workbook.View.ActiveTab == deletedSheetId)
		{
			_pck.Workbook.Worksheets[_pck._worksheetAdd].View.TabSelected = true;
		}
	}

	private void DeleteRelationsAndParts(Packaging.ZipPackagePart part)
	{
		var rels = part.GetRelationships().ToList();
		for (var i = 0; i < rels.Count; i++)
		{
			var rel = rels[i];
			if (rel.RelationshipType != ExcelPackage.schemaImage)
			{
				DeleteRelationsAndParts(_pck.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri)));
			}

			part.DeleteRelationship(rel.Id);
		}

		_pck.Package.DeletePart(part.Uri);
	}

	/// <summary>
	/// Deletes a worksheet from the collection
	/// </summary>
	/// <param name="name">The name of the worksheet in the workbook</param>
	public void Delete(string name)
	{
		var sheet = this[name];
		if (sheet == null)
		{
			throw new ArgumentException(string.Format("Could not find worksheet to delete '{0}'", name));
		}

		Delete(sheet.PositionID);
	}
	/// <summary>
	/// Delete a worksheet from the collection
	/// </summary>
	/// <param name="Worksheet">The worksheet to delete</param>
	public void Delete(ExcelWorksheet Worksheet)
	{
		if (Worksheet.PositionID <= _worksheets.Count && Worksheet == _worksheets[Worksheet.PositionID])
		{
			Delete(Worksheet.PositionID);
		}
		else
		{
			throw new ArgumentException("Worksheet is not in the collection.");
		}
	}
	internal void ReindexWorksheetDictionary()
	{
		var index = _pck._worksheetAdd;
		var worksheets = new Dictionary<int, ExcelWorksheet>();
		foreach (var entry in _worksheets)
		{
			entry.Value.PositionID = index;
			worksheets.Add(index++, entry.Value);
		}

		_worksheets = worksheets;
	}
}

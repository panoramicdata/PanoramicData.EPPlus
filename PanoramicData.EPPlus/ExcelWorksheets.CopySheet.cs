using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	private static void CopySheetNames(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		foreach (var name in Copy.Names)
		{
			ExcelNamedRange newName;
			if (!name.IsName)
			{
				if (name.WorkSheet == Copy.Name)
				{
					newName = added.Names.Add(name.Name, added.Cells[name.FirstAddress]);
				}
				else
				{
					newName = added.Names.Add(name.Name, added.Workbook.Worksheets[name.WorkSheet].Cells[name.FirstAddress]);
				}
			}
			else if (!string.IsNullOrEmpty(name.NameFormula))
			{
				newName = added.Names.AddFormula(name.Name, name.Formula);
			}
			else
			{
				newName = added.Names.AddValue(name.Name, name.Value);
			}

			newName.NameComment = name.NameComment;
		}
	}

	private void CopyTable(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		var prevName = "";
		//First copy the table XML
		foreach (var tbl in Copy.Tables)
		{
			var xml = tbl.TableXml.OuterXml;
			string name;
			if (prevName == "")
			{
				name = Copy.Tables.GetNewTableName();
			}
			else
			{
				var ix = int.Parse(prevName[5..]) + 1;
				name = string.Format("Table{0}", ix);
				while (_pck.Workbook.ExistsPivotTableName(name))
				{
					name = string.Format("Table{0}", ++ix);
				}
			}
			//ensure the _nextTableID value has been initialized - Pull request by WillR
			_pck.Workbook.ReadAllTables();

			var Id = _pck.Workbook._nextTableID++;
			prevName = name;
			XmlDocument xmlDoc = new();
			xmlDoc.LoadXml(xml);
			xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
			xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
			xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
			xml = xmlDoc.OuterXml;

			var uriTbl = GetNewUri(_pck.Package, "/xl/tables/table{0}.xml", ref Id);
			if (_pck.Workbook._nextTableID < Id)
			{
				_pck.Workbook._nextTableID = Id;
			}

			var part = _pck.Package.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", _pck.Compression);
			StreamWriter streamTbl = new(part.GetStream(FileMode.Create, FileAccess.Write));
			streamTbl.Write(xml);
			streamTbl.Flush();

			//create the relationship and add the ID to the worksheet xml.
			var rel = added.Part.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");

			if (tbl.RelationshipID == null)
			{
				var topNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
				if (topNode == null)
				{
					added.CreateNode("d:tableParts");
					topNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
				}

				var elem = added.WorksheetXml.CreateElement("tablePart", ExcelPackage.schemaMain);
				topNode.AppendChild(elem);
				elem.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
			}
			else
			{
				XmlAttribute relAtt;
				relAtt = added.WorksheetXml.SelectSingleNode(string.Format("//d:tableParts/d:tablePart/@r:id[.='{0}']", tbl.RelationshipID), tbl.NameSpaceManager) as XmlAttribute;
				relAtt.Value = rel.Id;
			}
		}
	}
	private void CopyPivotTable(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		var prevName = "";
		foreach (var tbl in Copy.PivotTables)
		{
			var xml = tbl.PivotTableXml.OuterXml;

			string name;
			if (Copy.Workbook == added.Workbook || added.PivotTables._pivotTableNames.ContainsKey(tbl.Name))
			{
				if (prevName == "")
				{
					name = added.PivotTables.GetNewTableName();
				}
				else
				{
					var ix = int.Parse(prevName[10..]) + 1;
					name = string.Format("PivotTable{0}", ix);
					while (_pck.Workbook.ExistsPivotTableName(name))
					{
						name = string.Format("PivotTable{0}", ++ix);
					}
				}
			}
			else
			{
				name = tbl.Name;
			}

			prevName = name;
			XmlDocument xmlDoc = new();
			//TODO: Fix save pivottable here
			xmlDoc.LoadXml(xml);
			xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@name", tbl.NameSpaceManager).Value = name;
			var cacheId = tbl.CacheID;
			if (!added.Workbook.ExistsPivotCache(tbl.CacheID, ref cacheId))
			{
				xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@cacheId", tbl.NameSpaceManager).Value = cacheId.ToString();

			}

			xml = xmlDoc.OuterXml;

			var Id = _pck.Workbook._nextPivotTableID++;
			var uriTbl = GetNewUri(_pck.Package, "/xl/pivotTables/pivotTable{0}.xml", ref Id);
			if (_pck.Workbook._nextPivotTableID < Id)
			{
				_pck.Workbook._nextPivotTableID = Id;
			}

			var partTbl = _pck.Package.CreatePart(uriTbl, ExcelPackage.schemaPivotTable, _pck.Compression);
			StreamWriter streamTbl = new(partTbl.GetStream(FileMode.Create, FileAccess.Write));
			streamTbl.Write(xml);
			streamTbl.Flush();

			xml = tbl.CacheDefinition.CacheDefinitionXml.OuterXml;
			var uriCd = GetNewUri(_pck.Package, "/xl/pivotCache/pivotcachedefinition{0}.xml", ref Id);
			var partCd = _pck.Package.CreatePart(uriCd, ExcelPackage.schemaPivotCacheDefinition, _pck.Compression);
			StreamWriter streamCd = new(partCd.GetStream(FileMode.Create, FileAccess.Write));
			streamCd.Write(xml);
			streamCd.Flush();

			added.Workbook.AddPivotTable(cacheId.ToString(), uriCd);

			xml = "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />";
			var uriRec = new Uri(string.Format("/xl/pivotCache/pivotCacheRecords{0}.xml", Id), UriKind.Relative);
			while (_pck.Package.PartExists(uriRec))
			{
				uriRec = new Uri(string.Format("/xl/pivotCache/pivotCacheRecords{0}.xml", ++Id), UriKind.Relative);
			}

			var partRec = _pck.Package.CreatePart(uriRec, ExcelPackage.schemaPivotCacheRecords, _pck.Compression);
			StreamWriter streamRec = new(partRec.GetStream(FileMode.Create, FileAccess.Write));
			streamRec.Write(xml);
			streamRec.Flush();

			//create the relationship and add the ID to the worksheet xml.
			added.Part.CreateRelationship(UriHelper.ResolvePartUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
			partTbl.CreateRelationship(UriHelper.ResolvePartUri(tbl.Relationship.SourceUri, uriCd), tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);
			partCd.CreateRelationship(UriHelper.ResolvePartUri(uriCd, uriRec), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");

		}

		added._pivotTables = null;   //Reset collection so it's reloaded when accessing the collection next time.
	}
	private void CopyHeaderFooterPictures(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		if (Copy.TopNode != null && Copy.TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager) == null)
		{
			return;
		}
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

	private static void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
	{
		to.LeftAlignedText = from.LeftAlignedText;
		to.CenteredText = from.CenteredText;
		to.RightAlignedText = from.RightAlignedText;
	}
}

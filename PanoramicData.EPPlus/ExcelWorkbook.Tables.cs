using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class ExcelWorkbook
{
	internal bool ExistsPivotCache(int cacheID, ref int newID)
	{
		newID = cacheID;
		var ret = true;
		foreach (var ws in Worksheets)
		{
			foreach (var pt in ws.PivotTables)
			{
				if (pt.CacheID == cacheID)
				{
					ret = false;
				}

				if (pt.CacheID >= newID)
				{
					newID = pt.CacheID + 1;
				}
			}
		}

		if (ret)
		{
			newID = cacheID;   //Not Found, return same ID
		}

		return ret;
	}

	internal bool ExistsTableName(string Name)
	{
		foreach (var ws in Worksheets)
		{
			if (ws.Tables._tableNames.ContainsKey(Name))
			{
				return true;
			}
		}

		return false;
	}
	internal bool ExistsPivotTableName(string Name)
	{
		foreach (var ws in Worksheets)
		{
			if (ws.PivotTables._pivotTableNames.ContainsKey(Name))
			{
				return true;
			}
		}

		return false;
	}
	internal void GetExternalReferences()
	{
		var nl = WorkbookXml.SelectNodes("//d:externalReferences/d:externalReference", NameSpaceManager);
		if (nl != null)
		{
			foreach (XmlElement elem in nl)
			{
				var rID = elem.GetAttribute("r:id");
				var rel = Part.GetRelationship(rID);
				var part = _package.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
				XmlDocument xmlExtRef = new();
				LoadXmlSafe(xmlExtRef, part.GetStream());

				if (xmlExtRef.SelectSingleNode("//d:externalBook", NameSpaceManager) is XmlElement book)
				{
					var rId_ExtRef = book.GetAttribute("r:id");
					var rel_extRef = part.GetRelationship(rId_ExtRef);
					if (rel_extRef != null)
					{
						_externalReferences.Add(rel_extRef.TargetUri.OriginalString);
					}

				}
			}
		}
	}

	internal void ReadAllTables()
	{
		if (_nextTableID > 0)
		{
			return;
		}

		_nextTableID = 1;
		_nextPivotTableID = 1;
		foreach (var ws in Worksheets)
		{
			if (ws is not ExcelChartsheet) //Fixes 15273. Chartsheets should be ignored.
			{
				foreach (var tbl in ws.Tables)
				{
					if (tbl.Id >= _nextTableID)
					{
						_nextTableID = tbl.Id + 1;
					}
				}

				foreach (var pt in ws.PivotTables)
				{
					if (pt.CacheID >= _nextPivotTableID)
					{
						_nextPivotTableID = pt.CacheID + 1;
					}
				}
			}
		}
	}
}

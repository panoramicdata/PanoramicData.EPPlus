using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void SavePivotTables()
	{
		foreach (var pt in PivotTables)
		{
			if (pt.DataFields.Count > 1)
			{
				AddDataFieldsPlaceholder(pt);
			}

			//Rewrite the pivottable address again if any rows or columns have been inserted or deleted
			pt.SetXmlNodeString("d:location/@ref", pt.Address.Address);
			var r = pt.CacheDefinition.SourceRange;
			if (r != null)  //Source does not exist
			{
				UpdatePivotCacheSource(pt);
			}

			pt.PivotTableXml.Save(pt.Part.GetStream(FileMode.Create));
			pt.CacheDefinition.CacheDefinitionXml.Save(pt.CacheDefinition.Part.GetStream(FileMode.Create));
		}
	}

	private static void AddDataFieldsPlaceholder(ExcelPivotTable pt)
	{
		XmlElement parentNode;
		if (pt.DataOnRows == true)
		{
			parentNode = pt.PivotTableXml.SelectSingleNode("//d:rowFields", pt.NameSpaceManager) as XmlElement;
			if (parentNode == null)
			{
				pt.CreateNode("d:rowFields");
				parentNode = pt.PivotTableXml.SelectSingleNode("//d:rowFields", pt.NameSpaceManager) as XmlElement;
			}
		}
		else
		{
			parentNode = pt.PivotTableXml.SelectSingleNode("//d:colFields", pt.NameSpaceManager) as XmlElement;
			if (parentNode == null)
			{
				pt.CreateNode("d:colFields");
				parentNode = pt.PivotTableXml.SelectSingleNode("//d:colFields", pt.NameSpaceManager) as XmlElement;
			}
		}

		if (parentNode.SelectSingleNode("d:field[@ x= \"-2\"]", pt.NameSpaceManager) == null)
		{
			var fieldNode = pt.PivotTableXml.CreateElement("field", ExcelPackage.schemaMain);
			fieldNode.SetAttribute("x", "-2");
			parentNode.AppendChild(fieldNode);
		}
	}

	private void UpdatePivotCacheSource(ExcelPivotTable pt)
	{
		var t = SetPivotCacheSourceReference(pt);

		var fields =
			pt.CacheDefinition.CacheDefinitionXml.SelectNodes(
				"d:pivotCacheDefinition/d:cacheFields/d:cacheField", NameSpaceManager);
		if (fields != null)
		{
			SetPivotCacheFieldNames(pt, fields, t);
			SetPivotDataFieldNames(pt);
		}
	}

	private static ExcelTable SetPivotCacheSourceReference(ExcelPivotTable pt)
	{
		ExcelTable t = null;
		if (pt.CacheDefinition.SourceRange.IsName)
		{
			//Named range, set name
			pt.CacheDefinition.DeleteNode(ExcelPivotCacheDefinition._sourceAddressPath); //Remove any address if previously set.
			pt.CacheDefinition.SetXmlNodeString(ExcelPivotCacheDefinition._sourceNamePath, ((ExcelNamedRange)pt.CacheDefinition.SourceRange).Name);
		}
		else
		{
			t = ExcelTableCollection.GetFromRange(pt.CacheDefinition.SourceRange);
			if (t == null)
			{
				//Address
				pt.CacheDefinition.DeleteNode(ExcelPivotCacheDefinition._sourceNamePath); //Remove any name or table if previously set.
				pt.CacheDefinition.SetXmlNodeString(ExcelPivotCacheDefinition._sourceAddressPath, pt.CacheDefinition.SourceRange.Address);
			}
			else
			{
				//Table, set name
				pt.CacheDefinition.DeleteNode(ExcelPivotCacheDefinition._sourceAddressPath); //Remove any address if previously set.
				pt.CacheDefinition.SetXmlNodeString(ExcelPivotCacheDefinition._sourceNamePath, t.Name);
			}
		}

		return t;
	}

}

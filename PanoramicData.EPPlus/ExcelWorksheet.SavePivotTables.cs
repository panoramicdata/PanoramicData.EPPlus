using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System.Globalization;
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

			//Rewrite the pivottable address again if any rows or columns have been inserted or deleted
			pt.SetXmlNodeString("d:location/@ref", pt.Address.Address);
			var r = pt.CacheDefinition.SourceRange;
			if (r != null)  //Source does not exist
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

				var fields =
					pt.CacheDefinition.CacheDefinitionXml.SelectNodes(
						"d:pivotCacheDefinition/d:cacheFields/d:cacheField", NameSpaceManager);
				var ix = 0;
				if (fields != null)
				{
					var flds = new HashSet<string>();
					foreach (XmlElement node in fields)
					{
						if (ix >= pt.CacheDefinition.SourceRange.Columns)
						{
							break;
						}

						var fldName = node.GetAttribute("name");                        //Fixes issue 15295 dup name error
						if (string.IsNullOrEmpty(fldName))
						{
							fldName = (t == null
								? pt.CacheDefinition.SourceRange.Offset(0, ix++, 1, 1).Value.ToString()
								: t.Columns[ix++].Name);
						}

						if (flds.Contains(fldName))
						{
							fldName = GetNewName(flds, fldName);
						}

						flds.Add(fldName);
						node.SetAttribute("name", fldName);
					}

					foreach (var df in pt.DataFields)
					{
						if (string.IsNullOrEmpty(df.Name))
						{
							string name;
							if (df.Function == DataFieldFunctions.None)
							{
								name = df.Field.Name; //Name must be set or Excel will crash on rename.                                
							}
							else
							{
								name = df.Function.ToString() + " of " + df.Field.Name; //Name must be set or Excel will crash on rename.
							}
							//Make sure name is unique
							var newName = name;
							var i = 2;
							while (pt.DataFields.ExistsDfName(newName, df))
							{
								newName = name + (i++).ToString(CultureInfo.InvariantCulture);
							}

							df.Name = newName;
						}
					}
				}
			}

			pt.PivotTableXml.Save(pt.Part.GetStream(FileMode.Create));
			pt.CacheDefinition.CacheDefinitionXml.Save(pt.CacheDefinition.Part.GetStream(FileMode.Create));
		}
	}

	private static string GetNewName(HashSet<string> flds, string fldName)
	{
		var ix = 2;
		while (flds.Contains(fldName + ix.ToString(CultureInfo.InvariantCulture)))
		{
			ix++;
		}

		return fldName + ix.ToString(CultureInfo.InvariantCulture);
	}
}

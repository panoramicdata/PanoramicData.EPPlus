using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private static void SetPivotCacheFieldNames(ExcelPivotTable pt, XmlNodeList fields, ExcelTable t)
	{
		var ix = 0;
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
	}

	private static void SetPivotDataFieldNames(ExcelPivotTable pt)
	{
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

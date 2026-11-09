using OfficeOpenXml.Compatibility;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class ExcelWorkbook
{
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	private void UpdateDefinedNamesXml()
	{
		try
		{
			var top = WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
			if (!ExistsNames())
			{
				if (top != null)
				{
					TopNode.RemoveChild(top);
				}

				return;
			}
			else
			{
				if (top == null)
				{
					CreateNode("d:definedNames");
					top = WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
				}
				else
				{
					top.RemoveAll();
				}

				foreach (var name in _names)
				{

					var elem = WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
					top.AppendChild(elem);
					elem.SetAttribute("name", name.Name);
					if (name.IsNameHidden)
					{
						elem.SetAttribute("hidden", "1");
					}

					if (!string.IsNullOrEmpty(name.NameComment))
					{
						elem.SetAttribute("comment", name.NameComment);
					}

					SetNameElement(name, elem);
				}
			}

			foreach (var ws in _worksheets)
			{
				if (ws is not ExcelChartsheet)
				{
					foreach (var name in ws.Names)
					{
						var elem = WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
						top.AppendChild(elem);
						elem.SetAttribute("name", name.Name);
						elem.SetAttribute("localSheetId", name.LocalSheetId.ToString());
						if (name.IsNameHidden)
						{
							elem.SetAttribute("hidden", "1");
						}

						if (!string.IsNullOrEmpty(name.NameComment))
						{
							elem.SetAttribute("comment", name.NameComment);
						}

						SetNameElement(name, elem);
					}
				}
			}
		}
		catch (Exception ex)
		{
			throw new Exception("Internal error updating named ranges ", ex);
		}
	}

	private static void SetNameElement(ExcelNamedRange name, XmlElement elem)
	{
		if (name.IsName)
		{
			if (string.IsNullOrEmpty(name.NameFormula))
			{
				if (TypeCompat.IsPrimitive(name.NameValue) || name.NameValue is double || name.NameValue is decimal)
				{
					elem.InnerText = Convert.ToDouble(name.NameValue, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
				}
				else if (name.NameValue is DateTime)
				{
					elem.InnerText = ((DateTime)name.NameValue).ToOADate().ToString(CultureInfo.InvariantCulture);
				}
				else
				{
					elem.InnerText = "\"" + name.NameValue.ToString() + "\"";
				}
			}
			else
			{
				elem.InnerText = name.NameFormula;
			}
		}
		else
		{
			elem.InnerText = name.FullAddressAbsolute;
		}
	}
	/// <summary>
	/// Is their any names in the workbook or in the sheets.
	/// </summary>
	/// <returns>?</returns>
	private bool ExistsNames()
	{
		if (_names.Count == 0)
		{
			foreach (var ws in Worksheets)
			{
				if (ws is ExcelChartsheet)
				{
					continue;
				}

				if (ws.Names.Count > 0)
				{
					return true;
				}
			}
		}
		else
		{
			return true;
		}

		return false;
	}
}

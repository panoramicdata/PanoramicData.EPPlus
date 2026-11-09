using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class OfficeProperties
{
	/// <summary>
	/// Provides access to the XML document which holds the document's custom properties
	/// </summary>
	public XmlDocument CustomPropertiesXml
	{
		get
		{
			_xmlPropertiesCustom ??= GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
						ExcelPackage.schemaVt,
						ExcelPackage.schemaCustom),
					 _uriPropertiesCustom,
					 @"application/vnd.openxmlformats-officedocument.custom-properties+xml",
					 @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");

			return _xmlPropertiesCustom;
		}
	}

	/// <summary>
	/// Gets the value of a custom property
	/// </summary>
	/// <param name="propertyName">The name of the property</param>
	/// <returns>The current value of the property</returns>
	public object GetCustomPropertyValue(string propertyName)
	{
		var searchString = string.Format("ctp:Properties/ctp:property[@name='{0}']", propertyName);
		if (CustomPropertiesXml.SelectSingleNode(searchString, NameSpaceManager) is XmlElement node)
		{
			return ParseCustomPropertyValue(node.LastChild.LocalName, node.LastChild.InnerText);
		}
		else
		{
			return null;
		}
	}

	private static object ParseCustomPropertyValue(string localName, string value)
	{
		switch (localName)
		{
			case "filetime":
				DateTime dt;
				return DateTime.TryParse(value, out dt) ? dt : null;
			case "i4":
				int i;
				return int.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out i) ? i : null;
			case "r8":
				double d;
				return double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out d) ? d : null;
			case "bool":
				if (value == "true")
				{
					return true;
				}
				else
				{
					return value == "false" ? false : null;
				}
			default:
				return value;
		}
	}

	/// <summary>
	/// Allows you to set the value of a current custom property or create your own custom property.  
	/// </summary>
	/// <param name="propertyName">The name of the property</param>
	/// <param name="value">The value of the property</param>
	public void SetCustomPropertyValue(string propertyName, object value)
	{
		var allProps = CustomPropertiesXml.SelectSingleNode(@"ctp:Properties", NameSpaceManager);

		var prop = string.Format("ctp:Properties/ctp:property[@name='{0}']", propertyName);
		if (CustomPropertiesXml.SelectSingleNode(prop, NameSpaceManager) is not XmlElement node)
		{
			node = CreateCustomPropertyNode(allProps, propertyName);
		}
		else
		{
			while (node.ChildNodes.Count > 0)
			{
				node.RemoveChild(node.ChildNodes[0]);
			}
		}

		node.AppendChild(CreateCustomPropertyValueElement(value));
	}

	private XmlElement CreateCustomPropertyNode(XmlNode allProps, string propertyName)
	{
		int pid;
		var MaxNode = CustomPropertiesXml.SelectSingleNode("ctp:Properties/ctp:property[not(@pid <= preceding-sibling::ctp:property/@pid) and not(@pid <= following-sibling::ctp:property/@pid)]", NameSpaceManager);
		if (MaxNode == null)
		{
			pid = 2;
		}
		else
		{
			if (!int.TryParse(MaxNode.Attributes["pid"].Value, out pid))
			{
				pid = 2;
			}

			pid++;
		}

		var node = CustomPropertiesXml.CreateElement("property", ExcelPackage.schemaCustom);
		node.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
		node.SetAttribute("pid", pid.ToString());  // custom property pid
		node.SetAttribute("name", propertyName);

		allProps.AppendChild(node);
		return node;
	}

	private XmlElement CreateCustomPropertyValueElement(object value)
	{
		XmlElement valueElem;
		if (value is bool)
		{
			valueElem = CustomPropertiesXml.CreateElement("vt", "bool", ExcelPackage.schemaVt);
			valueElem.InnerText = value.ToString().ToLower(CultureInfo.InvariantCulture);
		}
		else if (value is DateTime dateTime)
		{
			valueElem = CustomPropertiesXml.CreateElement("vt", "filetime", ExcelPackage.schemaVt);
			valueElem.InnerText = dateTime.AddHours(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
		}
		else if (value is short or int)
		{
			valueElem = CustomPropertiesXml.CreateElement("vt", "i4", ExcelPackage.schemaVt);
			valueElem.InnerText = value.ToString();
		}
		else if (value is double or decimal or float or long)
		{
			valueElem = CustomPropertiesXml.CreateElement("vt", "r8", ExcelPackage.schemaVt);
			valueElem.InnerText = GetR8Text(value);
		}
		else
		{
			valueElem = CustomPropertiesXml.CreateElement("vt", "lpwstr", ExcelPackage.schemaVt);
			valueElem.InnerText = value.ToString();
		}

		return valueElem;
	}

	private static string GetR8Text(object value)
	{
		if (value is double d)
		{
			return d.ToString(CultureInfo.InvariantCulture);
		}
		else if (value is float f)
		{
			return f.ToString(CultureInfo.InvariantCulture);
		}
		else if (value is decimal m)
		{
			return m.ToString(CultureInfo.InvariantCulture);
		}
		else
		{
			return value.ToString();
		}
	}
}

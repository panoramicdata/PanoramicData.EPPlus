using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public abstract partial class XmlHelper
{
	internal bool ExistNode(string path) => TopNode != null && TopNode.SelectSingleNode(path, NameSpaceManager) != null;
	internal bool? GetXmlNodeBoolNullable(string path)
	{
		var value = GetXmlNodeString(path);
		return string.IsNullOrEmpty(value) ? null : GetXmlNodeBool(path);
	}
	internal bool GetXmlNodeBool(string path) => GetXmlNodeBool(path, false);
	internal bool GetXmlNodeBool(string path, bool blankValue)
	{
		var value = GetXmlNodeString(path);
		if (value == "1" || value == "-1" || value.Equals("true", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		else
		{
			return value == "" && blankValue;
		}
	}
	internal int GetXmlNodeInt(string path)
		=> int.TryParse(GetXmlNodeString(path), NumberStyles.Number, CultureInfo.InvariantCulture, out var i) ? i : int.MinValue;

	internal int? GetXmlNodeIntNull(string path)
	{
		var s = GetXmlNodeString(path);
		return s != "" && int.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var i) ? i : null;
	}

	internal decimal GetXmlNodeDecimal(string path)
		=> decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;

	internal decimal? GetXmlNodeDecimalNull(string path)
		=> decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : null;

	internal double? GetXmlNodeDoubleNull(string path)
	{
		var s = GetXmlNodeString(path);
		if (s == "")
		{
			return null;
		}
		else
		{
			return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;
		}
	}
	internal double GetXmlNodeDouble(string path)
	{
		var s = GetXmlNodeString(path);
		if (s == "")
		{
			return double.NaN;
		}
		else
		{
			return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : double.NaN;
		}
	}

	internal string GetXmlNodeString(XmlNode node, string path)
	{
		if (node == null)
		{
			return "";
		}

		var nameNode = node.SelectSingleNode(path, NameSpaceManager);

		if (nameNode != null)
		{
			return nameNode.NodeType == XmlNodeType.Attribute ? nameNode.Value != null ? nameNode.Value : "" : nameNode.InnerText;
		}
		else
		{
			return "";
		}
	}
	internal string GetXmlNodeString(string path) => GetXmlNodeString(TopNode, path);
}

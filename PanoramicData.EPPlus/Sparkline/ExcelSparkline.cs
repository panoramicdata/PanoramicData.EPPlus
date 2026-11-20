using System.Xml;

namespace OfficeOpenXml.Sparkline;

/// <summary>
/// Represents a single sparkline within the sparkline group
/// </summary>
public class ExcelSparkline : XmlHelper
{
	internal ExcelSparkline(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
	{
		SchemaNodeOrder = ["f", "sqref"];
	}
	const string _fPath = "xm:f";
	/// <summary>
	/// Sets the datarange
	/// </summary>
	internal void SetRangeAddress(ExcelAddressBase value)
	{
		if (value is ExcelNamedRange namedRange)
		{
			SetXmlNodeString(_fPath, namedRange.Name);
		}
		else
		{
			SetXmlNodeString(_fPath, value.FullAddress);
		}
	}

	/// <summary>
	/// Get the data range address.
	/// </summary>
	/// <param name="namedRangeCol">workbook or worksheet Names</param>
	/// <returns></returns>
	internal ExcelAddressBase GetRangeAddress(ExcelNamedRangeCollection namedRangeCol)
	{
		var addrOrName = GetXmlNodeString(_fPath);
		return namedRangeCol.ContainsKey(addrOrName) ? namedRangeCol[addrOrName] : new ExcelAddressBase(addrOrName);
	}

	const string _sqrefPath = "xm:sqref";
	/// <summary>
	/// Location of the sparkline
	/// </summary>
	public ExcelCellAddress Cell
	{
		get
		{
			return new ExcelCellAddress(GetXmlNodeString(_sqrefPath));
		}
		internal set
		{
			SetXmlNodeString("xm:sqref", value.Address);
		}
	}
	public override string ToString() => Cell.Address + ", " + GetXmlNodeString(_fPath);
}

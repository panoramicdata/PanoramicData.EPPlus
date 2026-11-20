namespace OfficeOpenXml.Table.PivotTable;

public partial class ExcelPivotTableField
{
	#region Show properties
	/// <summary>
	/// Indicates whether to show all items for this field
	/// </summary>
	public bool ShowAll
	{
		get
		{
			return GetXmlNodeBool("@showAll");
		}
		set
		{
			SetXmlNodeBool("@showAll", value);
		}
	}
	/// <summary>
	/// Indicates whether to hide drop down buttons on PivotField headers
	/// </summary>
	public bool ShowDropDowns
	{
		get
		{
			return GetXmlNodeBool("@showDropDowns");
		}
		set
		{
			SetXmlNodeBool("@showDropDowns", value);
		}
	}
	/// <summary>
	/// Indicates whether this hierarchy is omitted from the field list
	/// </summary>
	public bool ShowInFieldList
	{
		get
		{
			return GetXmlNodeBool("@showInFieldList");
		}
		set
		{
			SetXmlNodeBool("@showInFieldList", value);
		}
	}
	/// <summary>
	/// Indicates whether to show the property as a member caption
	/// </summary>
	public bool ShowAsCaption
	{
		get
		{
			return GetXmlNodeBool("@showPropAsCaption");
		}
		set
		{
			SetXmlNodeBool("@showPropAsCaption", value);
		}
	}
	/// <summary>
	/// Indicates whether to show the member property value in a PivotTable cell
	/// </summary>
	public bool ShowMemberPropertyInCell
	{
		get
		{
			return GetXmlNodeBool("@showPropCell");
		}
		set
		{
			SetXmlNodeBool("@showPropCell", value);
		}
	}
	/// <summary>
	/// Indicates whether to show the member property value in a tooltip on the appropriate PivotTable cells
	/// </summary>
	public bool ShowMemberPropertyToolTip
	{
		get
		{
			return GetXmlNodeBool("@showPropTip");
		}
		set
		{
			SetXmlNodeBool("@showPropTip", value);
		}
	}
	#endregion
}

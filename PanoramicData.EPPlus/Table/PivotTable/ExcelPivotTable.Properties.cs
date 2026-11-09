namespace OfficeOpenXml.Table.PivotTable;

public partial class ExcelPivotTable
{
	#region "Public Properties"

	/// <summary>
	/// The worksheet where the pivottable is located
	/// </summary>
	public ExcelWorksheet WorkSheet
	{
		get;
		set;
	}
	/// <summary>
	/// The location of the pivot table
	/// </summary>
	public ExcelAddressBase Address
	{
		get;
		internal set;
	}
	/// <summary>
	/// If multiple datafields are displayed in the row area or the column area
	/// </summary>
	public bool DataOnRows
	{
		get
		{
			return GetXmlNodeBool("@dataOnRows");
		}
		set
		{
			SetXmlNodeBool("@dataOnRows", value);
		}
	}
	/// <summary>
	/// if true apply legacy table autoformat number format properties.
	/// </summary>
	public bool ApplyNumberFormats
	{
		get
		{
			return GetXmlNodeBool("@applyNumberFormats");
		}
		set
		{
			SetXmlNodeBool("@applyNumberFormats", value);
		}
	}
	/// <summary>
	/// If true apply legacy table autoformat border properties
	/// </summary>
	public bool ApplyBorderFormats
	{
		get
		{
			return GetXmlNodeBool("@applyBorderFormats");
		}
		set
		{
			SetXmlNodeBool("@applyBorderFormats", value);
		}
	}
	/// <summary>
	/// If true apply legacy table autoformat font properties
	/// </summary>
	public bool ApplyFontFormats
	{
		get
		{
			return GetXmlNodeBool("@applyFontFormats");
		}
		set
		{
			SetXmlNodeBool("@applyFontFormats", value);
		}
	}
	/// <summary>
	/// If true apply legacy table autoformat pattern properties
	/// </summary>
	public bool ApplyPatternFormats
	{
		get
		{
			return GetXmlNodeBool("@applyPatternFormats");
		}
		set
		{
			SetXmlNodeBool("@applyPatternFormats", value);
		}
	}
	/// <summary>
	/// If true apply legacy table autoformat width/height properties.
	/// </summary>
	public bool ApplyWidthHeightFormats
	{
		get
		{
			return GetXmlNodeBool("@applyWidthHeightFormats");
		}
		set
		{
			SetXmlNodeBool("@applyWidthHeightFormats", value);
		}
	}
	/// <summary>
	/// Show member property information
	/// </summary>
	public bool ShowMemberPropertyTips
	{
		get
		{
			return GetXmlNodeBool("@showMemberPropertyTips");
		}
		set
		{
			SetXmlNodeBool("@showMemberPropertyTips", value);
		}
	}
	/// <summary>
	/// Show the drill indicators
	/// </summary>
	public bool ShowCalcMember
	{
		get
		{
			return GetXmlNodeBool("@showCalcMbrs");
		}
		set
		{
			SetXmlNodeBool("@showCalcMbrs", value);
		}
	}
	/// <summary>
	/// If the user is prevented from drilling down on a PivotItem or aggregate value
	/// </summary>
	public bool EnableDrill
	{
		get
		{
			return GetXmlNodeBool("@enableDrill", true);
		}
		set
		{
			SetXmlNodeBool("@enableDrill", value);
		}
	}
	/// <summary>
	/// Show the drill down buttons
	/// </summary>
	public bool ShowDrill
	{
		get
		{
			return GetXmlNodeBool("@showDrill", true);
		}
		set
		{
			SetXmlNodeBool("@showDrill", value);
		}
	}
	/// <summary>
	/// If the tooltips should be displayed for PivotTable data cells.
	/// </summary>
	public bool ShowDataTips
	{
		get
		{
			return GetXmlNodeBool("@showDataTips", true);
		}
		set
		{
			SetXmlNodeBool("@showDataTips", value, true);
		}
	}
	/// <summary>
	/// If the row and column titles from the PivotTable should be printed.
	/// </summary>
	public bool FieldPrintTitles
	{
		get
		{
			return GetXmlNodeBool("@fieldPrintTitles");
		}
		set
		{
			SetXmlNodeBool("@fieldPrintTitles", value);
		}
	}
	/// <summary>
	/// If the row and column titles from the PivotTable should be printed.
	/// </summary>
	public bool ItemPrintTitles
	{
		get
		{
			return GetXmlNodeBool("@itemPrintTitles");
		}
		set
		{
			SetXmlNodeBool("@itemPrintTitles", value);
		}
	}
	/// <summary>
	/// If the grand totals should be displayed for the PivotTable columns
	/// </summary>
	[Obsolete("Use correctly spelled property 'ColumnGrandTotals'")]
	public bool ColumGrandTotals
	{
		get
		{
			return ColumnGrandTotals;
		}
		set
		{
			ColumnGrandTotals = value;
		}
	}
	/// <summary>
	/// If the grand totals should be displayed for the PivotTable columns
	/// </summary>
	public bool ColumnGrandTotals
	{
		get
		{
			return GetXmlNodeBool("@colGrandTotals");
		}
		set
		{
			SetXmlNodeBool("@colGrandTotals", value);
		}
	}
	/// <summary>
	/// If the grand totals should be displayed for the PivotTable rows
	/// </summary>
	public bool RowGrandTotals
	{
		get
		{
			return GetXmlNodeBool("@rowGrandTotals");
		}
		set
		{
			SetXmlNodeBool("@rowGrandTotals", value);
		}
	}
	/// <summary>
	/// If the drill indicators expand collapse buttons should be printed.
	/// </summary>
	public bool PrintDrill
	{
		get
		{
			return GetXmlNodeBool("@printDrill");
		}
		set
		{
			SetXmlNodeBool("@printDrill", value);
		}
	}
	/// <summary>
	/// Indicates whether to show error messages in cells.
	/// </summary>
	public bool ShowError
	{
		get
		{
			return GetXmlNodeBool("@showError");
		}
		set
		{
			SetXmlNodeBool("@showError", value);
		}
	}
	/// <summary>
	/// The string to be displayed in cells that contain errors.
	/// </summary>
	public string ErrorCaption
	{
		get
		{
			return GetXmlNodeString("@errorCaption");
		}
		set
		{
			SetXmlNodeString("@errorCaption", value);
		}
	}
	/// <summary>
	/// Specifies the name of the value area field header in the PivotTable. 
	/// This caption is shown when the PivotTable when two or more fields are in the values area.
	/// </summary>
	public string DataCaption
	{
		get
		{
			return GetXmlNodeString("@dataCaption");
		}
		set
		{
			SetXmlNodeString("@dataCaption", value);
		}
	}
	/// <summary>
	/// Show field headers
	/// </summary>
	public bool ShowHeaders
	{
		get
		{
			return GetXmlNodeBool("@showHeaders");
		}
		set
		{
			SetXmlNodeBool("@showHeaders", value);
		}
	}
	/// <summary>
	/// The number of page fields to display before starting another row or column
	/// </summary>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Public API: callers may catch Exception from this setter, so the exception type is kept.")]
	public int PageWrap
	{
		get
		{
			return GetXmlNodeInt("@pageWrap");
		}
		set
		{
			if (value < 0)
			{
				throw new Exception("Value can't be negative");
			}

			SetXmlNodeString("@pageWrap", value.ToString());
		}
	}
	/// <summary>
	/// A boolean that indicates whether legacy auto formatting has been applied to the PivotTable view
	/// </summary>
	public bool UseAutoFormatting
	{
		get
		{
			return GetXmlNodeBool("@useAutoFormatting");
		}
		set
		{
			SetXmlNodeBool("@useAutoFormatting", value);
		}
	}
	/// <summary>
	/// A boolean that indicates whether the in-grid drop zones should be displayed at runtime, and whether classic layout is applied
	/// </summary>
	public bool GridDropZones
	{
		get
		{
			return GetXmlNodeBool("@gridDropZones");
		}
		set
		{
			SetXmlNodeBool("@gridDropZones", value);
		}
	}
	/// <summary>
	/// Specifies the indentation increment for compact axis and can be used to set the Report Layout to Compact Form
	/// </summary>
	public int Indent
	{
		get
		{
			return GetXmlNodeInt("@indent");
		}
		set
		{
			SetXmlNodeString("@indent", value.ToString());
		}
	}
	/// <summary>
	/// A boolean that indicates whether data fields in the PivotTable should be displayed in outline form
	/// </summary>
	public bool OutlineData
	{
		get
		{
			return GetXmlNodeBool("@outlineData");
		}
		set
		{
			SetXmlNodeBool("@outlineData", value);
		}
	}
	/// <summary>
	/// a boolean that indicates whether new fields should have their outline flag set to true
	/// </summary>
	public bool Outline
	{
		get
		{
			return GetXmlNodeBool("@outline");
		}
		set
		{
			SetXmlNodeBool("@outline", value);
		}
	}
	/// <summary>
	/// A boolean that indicates whether the fields of a PivotTable can have multiple filters set on them
	/// </summary>
	public bool MultipleFieldFilters
	{
		get
		{
			return GetXmlNodeBool("@multipleFieldFilters");
		}
		set
		{
			SetXmlNodeBool("@multipleFieldFilters", value);
		}
	}
	/// <summary>
	/// A boolean that indicates whether new fields should have their compact flag set to true
	/// </summary>
	public bool Compact
	{
		get
		{
			return GetXmlNodeBool("@compact");
		}
		set
		{
			SetXmlNodeBool("@compact", value);
		}
	}
	/// <summary>
	/// A boolean that indicates whether the field next to the data field in the PivotTable should be displayed in the same column of the spreadsheet
	/// </summary>
	public bool CompactData
	{
		get
		{
			return GetXmlNodeBool("@compactData");
		}
		set
		{
			SetXmlNodeBool("@compactData", value);
		}
	}
	/// <summary>
	/// Specifies the string to be displayed for grand totals.
	/// </summary>
	public string GrandTotalCaption
	{
		get
		{
			return GetXmlNodeString("@grandTotalCaption");
		}
		set
		{
			SetXmlNodeString("@grandTotalCaption", value);
		}
	}
	/// <summary>
	/// Specifies the string to be displayed in row header in compact mode.
	/// </summary>
	public string RowHeaderCaption
	{
		get
		{
			return GetXmlNodeString("@rowHeaderCaption");
		}
		set
		{
			SetXmlNodeString("@rowHeaderCaption", value);
		}
	}
	/// <summary>
	/// Specifies the string to be displayed in column header in compact mode.
	/// </summary>
	public string ColumnHeaderCaption
	{
		get
		{
			return GetXmlNodeString("@colHeaderCaption");
		}
		set
		{
			SetXmlNodeString("@colHeaderCaption", value);
		}
	}
	/// <summary>
	/// Specifies the string to be displayed in cells with no value
	/// </summary>
	public string MissingCaption
	{
		get
		{
			return GetXmlNodeString("@missingCaption");
		}
		set
		{
			SetXmlNodeString("@missingCaption", value);
		}
	}
	const string FIRSTHEADERROW_PATH = "d:location/@firstHeaderRow";
	/// <summary>
	/// Specifies the first row of the PivotTable header, relative to the top left cell in the ref value
	/// </summary>
	public int FirstHeaderRow
	{
		get
		{
			return GetXmlNodeInt(FIRSTHEADERROW_PATH);
		}
		set
		{
			SetXmlNodeString(FIRSTHEADERROW_PATH, value.ToString());
		}
	}
	const string FIRSTDATAROW_PATH = "d:location/@firstDataRow";
	/// <summary>
	/// Specifies the first column of the PivotTable data, relative to the top left cell in the ref value
	/// </summary>
	public int FirstDataRow
	{
		get
		{
			return GetXmlNodeInt(FIRSTDATAROW_PATH);
		}
		set
		{
			SetXmlNodeString(FIRSTDATAROW_PATH, value.ToString());
		}
	}
	const string FIRSTDATACOL_PATH = "d:location/@firstDataCol";
	/// <summary>
	/// Specifies the first column of the PivotTable data, relative to the top left cell in the ref value
	/// </summary>
	public int FirstDataCol
	{
		get
		{
			return GetXmlNodeInt(FIRSTDATACOL_PATH);
		}
		set
		{
			SetXmlNodeString(FIRSTDATACOL_PATH, value.ToString());
		}
	}
	#endregion
}

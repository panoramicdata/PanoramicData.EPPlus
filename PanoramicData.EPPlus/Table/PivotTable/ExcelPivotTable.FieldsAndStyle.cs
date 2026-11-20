namespace OfficeOpenXml.Table.PivotTable;

public partial class ExcelPivotTable
{
	ExcelPivotTableFieldCollection _fields;
	/// <summary>
	/// The fields in the table 
	/// </summary>
	public ExcelPivotTableFieldCollection Fields
	{
		get
		{
			_fields ??= new ExcelPivotTableFieldCollection(this, "");

			return _fields;
		}
	}
	ExcelPivotTableRowColumnFieldCollection _rowFields;
	/// <summary>
	/// Row label fields 
	/// </summary>
	public ExcelPivotTableRowColumnFieldCollection RowFields
	{
		get
		{
			_rowFields ??= new ExcelPivotTableRowColumnFieldCollection(this, "rowFields");

			return _rowFields;
		}
	}
	ExcelPivotTableRowColumnFieldCollection _columnFields;
	/// <summary>
	/// Column label fields 
	/// </summary>
	public ExcelPivotTableRowColumnFieldCollection ColumnFields
	{
		get
		{
			_columnFields ??= new ExcelPivotTableRowColumnFieldCollection(this, "colFields");

			return _columnFields;
		}
	}
	ExcelPivotTableDataFieldCollection _dataFields;
	/// <summary>
	/// Value fields 
	/// </summary>
	public ExcelPivotTableDataFieldCollection DataFields
	{
		get
		{
			_dataFields ??= new ExcelPivotTableDataFieldCollection(this);

			return _dataFields;
		}
	}
	ExcelPivotTableRowColumnFieldCollection _pageFields;
	/// <summary>
	/// Report filter fields
	/// </summary>
	public ExcelPivotTableRowColumnFieldCollection PageFields
	{
		get
		{
			_pageFields ??= new ExcelPivotTableRowColumnFieldCollection(this, "pageFields");

			return _pageFields;
		}
	}
	const string STYLENAME_PATH = "d:pivotTableStyleInfo/@name";
	/// <summary>
	/// Pivot style name. Used for custom styles
	/// </summary>
	public string StyleName
	{
		get
		{
			return GetXmlNodeString(STYLENAME_PATH);
		}
		set
		{
			(_tableStyle, value) = TableNameHelper.ParseStyleName(value, "PivotStyle");
			SetXmlNodeString(STYLENAME_PATH, value, true);
		}
	}
	TableStyles _tableStyle = TableStyles.Medium6;
	/// <summary>
	/// The table style. If this property is cusom the style from the StyleName propery is used.
	/// </summary>
	public TableStyles TableStyle
	{
		get
		{
			return _tableStyle;
		}
		set
		{
			_tableStyle = value;
			if (value != TableStyles.Custom)
			{
				SetXmlNodeString(STYLENAME_PATH, "PivotStyle" + value.ToString());
			}
		}
	}
}

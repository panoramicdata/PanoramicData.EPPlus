namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// Defines the axis for a PivotTable
/// </summary>
public enum ePivotFieldAxis
{
	/// <summary>
	/// None
	/// </summary>
	None = -1,
	/// <summary>
	/// Column axis
	/// </summary>
	Column,
	/// <summary>
	/// Page axis (Include Count Filter) 
	/// 
	/// </summary>
	Page,
	/// <summary>
	/// Row axis
	/// </summary>
	Row,
	/// <summary>
	/// Values axis
	/// </summary>
	Values
}
/// <summary>
/// Build-in table row functions
/// </summary>
public enum DataFieldFunctions
{
	Average,
	Count,
	CountNums,
	Max,
	Min,
	Product,
	None,
	StdDev,
	StdDevP,
	Sum,
	Var,
	VarP
}
/// <summary>
/// Defines the data formats for a field in the PivotTable
/// </summary>
public enum eShowDataAs
{
	/// <summary>
	/// Indicates the field is shown as the "difference from" a value.
	/// </summary>
	Difference,
	/// <summary>
	/// Indicates the field is shown as the "index.
	/// </summary>
	Index,
	/// <summary>
	/// Indicates that the field is shown as its normal datatype.
	/// </summary>
	Normal,
	/// <summary>
	/// /Indicates the field is show as the "percentage of" a value
	/// </summary>
	Percent,
	/// <summary>
	/// Indicates the field is shown as the "percentage difference from" a value.
	/// </summary>
	PercentDiff,
	/// <summary>
	/// Indicates the field is shown as the percentage of column.
	/// </summary>
	PercentOfCol,
	/// <summary>
	/// Indicates the field is shown as the percentage of row
	/// </summary>
	PercentOfRow,
	/// <summary>
	/// Indicates the field is shown as percentage of total.
	/// </summary>
	PercentOfTotal,
	/// <summary>
	/// Indicates the field is shown as running total in the table.
	/// </summary>
	RunTotal,
}
/// <summary>
/// Built-in subtotal functions
/// </summary>
[Flags]
public enum eSubTotalFunctions
{
	None = 1,
	Count = 2,
	CountA = 4,
	Avg = 8,
	Default = 16,
	Min = 32,
	Max = 64,
	Product = 128,
	StdDev = 256,
	StdDevP = 512,
	Sum = 1024,
	Var = 2048,
	VarP = 4096
}
/// <summary>
/// Data grouping
/// </summary>
[Flags]
public enum eDateGroupBy
{
	Years = 1,
	Quarters = 2,
	Months = 4,
	Days = 8,
	Hours = 16,
	Minutes = 32,
	Seconds = 64
}
/// <summary>
/// Sorting
/// </summary>
public enum eSortType
{
	None,
	Ascending,
	Descending
}

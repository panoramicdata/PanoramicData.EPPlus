using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Dimension address for the worksheet. 
	/// Top left cell to Bottom right.
	/// If the worksheet has no cells, null is returned
	/// </summary>
	public ExcelAddressBase? Dimension
	{
		get
		{
			CheckSheetType();
			if (_values.GetDimension(out var fromRow, out var fromCol, out var toRow, out var toCol))
			{
				var addr = new ExcelAddressBase(fromRow, fromCol, toRow, toCol)
				{
					_ws = Name
				};
				return addr;
			}
			else
			{
				return null;
			}
		}
	}
	ExcelSheetProtection? _protection;
	/// <summary>
	/// Access to sheet protection properties
	/// </summary>
	public ExcelSheetProtection Protection
	{
		get
		{
			_protection ??= new ExcelSheetProtection(NameSpaceManager, TopNode, this);

			return _protection;
		}
	}

	private ExcelProtectedRangeCollection _protectedRanges;
	public ExcelProtectedRangeCollection ProtectedRanges
	{
		get
		{
			_protectedRanges ??= new ExcelProtectedRangeCollection(NameSpaceManager, TopNode, this);
			return _protectedRanges;
		}
	}

	ExcelDrawings? _drawings;
	/// <summary>
	/// Collection of drawing-objects like shapes, images and charts
	/// </summary>
	public ExcelDrawings Drawings
	{
		get
		{
			_drawings ??= new ExcelDrawings(_package, this);

			return _drawings;
		}
	}

	ExcelSparklineGroupCollection? _sparklineGroups;
	/// <summary>
	/// Collection of Sparkline-objects. 
	/// Sparklines are small in-cell charts.
	/// </summary>
	public ExcelSparklineGroupCollection SparklineGroups
	{
		get
		{
			_sparklineGroups ??= new ExcelSparklineGroupCollection(this);

			return _sparklineGroups;
		}
	}

	ExcelTableCollection? _tables;
	/// <summary>
	/// Tables defined in the worksheet.
	/// </summary>
	public ExcelTableCollection Tables
	{
		get
		{
			CheckSheetType();
			if (Workbook._nextTableID == int.MinValue)
			{
				Workbook.ReadAllTables();
			}

			_tables ??= new ExcelTableCollection(this);

			return _tables;
		}
	}
	internal ExcelPivotTableCollection? _pivotTables;
	/// <summary>
	/// Pivottables defined in the worksheet.
	/// </summary>
	public ExcelPivotTableCollection PivotTables
	{
		get
		{
			CheckSheetType();
			if (_pivotTables == null)
			{
				if (Workbook._nextPivotTableID == int.MinValue)
				{
					Workbook.ReadAllTables();
				}

				_pivotTables = new ExcelPivotTableCollection(this);
			}

			return _pivotTables;
		}
	}
	private ExcelConditionalFormattingCollection? _conditionalFormatting;
	/// <summary>
	/// ConditionalFormatting defined in the worksheet. Use the Add methods to create ConditionalFormatting and add them to the worksheet. Then
	/// set the properties on the instance returned.
	/// </summary>
	/// <seealso cref="ExcelConditionalFormattingCollection"/>
	public ExcelConditionalFormattingCollection ConditionalFormatting
	{
		get
		{
			CheckSheetType();
			_conditionalFormatting ??= new ExcelConditionalFormattingCollection(this);

			return _conditionalFormatting;
		}
	}
	private ExcelDataValidationCollection? _dataValidation;
	/// <summary>
	/// DataValidation defined in the worksheet. Use the Add methods to create DataValidations and add them to the worksheet. Then
	/// set the properties on the instance returned.
	/// </summary>
	/// <seealso cref="ExcelDataValidationCollection"/>
	public ExcelDataValidationCollection DataValidations
	{
		get
		{
			CheckSheetType();
			_dataValidation ??= new ExcelDataValidationCollection(this);

			return _dataValidation;
		}
	}

	internal void ClearValidations() => _dataValidation = null;

	ExcelBackgroundImage? _backgroundImage;
	/// <summary>
	/// An image displayed as the background of the worksheet.
	/// </summary>
	public ExcelBackgroundImage BackgroundImage
	{
		get
		{
			_backgroundImage ??= new ExcelBackgroundImage(NameSpaceManager, TopNode, this);

			return _backgroundImage;
		}
	}
	/// <summary>
	/// Returns the style ID given a style name.  
	/// The style ID will be created if not found, but only if the style name exists!
	/// </summary>
	/// <param name="StyleName"></param>
	/// <returns></returns>
	internal int GetStyleID(string StyleName)
	{
		ExcelNamedStyleXml namedStyle = null;
		Workbook.Styles.NamedStyles.FindByID(StyleName, ref namedStyle);
		if (namedStyle.XfId == int.MinValue)
		{
			namedStyle.XfId = Workbook.Styles.CellXfs.FindIndexByID(namedStyle.Style.Id);
		}

		return namedStyle.XfId;
	}
}

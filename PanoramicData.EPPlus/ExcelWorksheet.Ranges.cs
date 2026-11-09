using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// A reference to the header and footer class which allows you to 
	/// set the header and footer for all odd, even and first pages of the worksheet
	/// </summary>
	/// <remarks>
	/// To format the text you can use the following format
	/// <list type="table">
	/// <listheader><term>Prefix</term><description>Description</description></listheader>
	/// <item><term>&amp;U</term><description>Underlined</description></item>
	/// <item><term>&amp;E</term><description>Double Underline</description></item>
	/// <item><term>&amp;K:xxxxxx</term><description>Color. ex &amp;K:FF0000 for red</description></item>
	/// <item><term>&amp;"Font,Regular Bold Italic"</term><description>Changes the font. Regular or Bold or Italic or Bold Italic can be used. ex &amp;"Arial,Bold Italic"</description></item>
	/// <item><term>&amp;nn</term><description>Change font size. nn is an integer. ex &amp;24</description></item>
	/// <item><term>&amp;G</term><description>Placeholder for images. Images can not be added by the library, but its possible to use in a template.</description></item>
	/// </list>
	/// </remarks>
	public ExcelHeaderFooter HeaderFooter
	{
		get
		{
			if (_headerFooter == null)
			{
				var headerFooterNode = TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager);
				headerFooterNode ??= CreateNode("d:headerFooter");
				_headerFooter = new ExcelHeaderFooter(NameSpaceManager, headerFooterNode, this);
			}

			return _headerFooter;
		}
	}

	/// <summary>
	/// Printer settings
	/// </summary>
	public ExcelPrinterSettings PrinterSettings
	{
		get
		{
			var ps = new ExcelPrinterSettings(NameSpaceManager, TopNode, this)
			{
				SchemaNodeOrder = SchemaNodeOrder
			};
			return ps;
		}
	}



	/// <summary>
	/// Provides access to a range of cells
	/// </summary>  
	public ExcelRange Cells
	{
		get
		{
			CheckSheetType();
			return new ExcelRange(this, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
		}
	}
	/// <summary>
	/// Provides access to the selected range of cells
	/// </summary>  
	public ExcelRange SelectedRange
	{
		get
		{
			CheckSheetType();
			return new ExcelRange(this, View.SelectedRange);
		}
	}
	readonly MergeCellsCollection _mergedCells = [];
	/// <summary>
	/// Addresses to merged ranges
	/// </summary>
	public MergeCellsCollection MergedCells
	{
		get
		{
			CheckSheetType();
			return _mergedCells;
		}
	}
	/// <summary>
	/// Provides access to an individual row within the worksheet so you can set its properties.
	/// </summary>
	/// <param name="row">The row number in the worksheet</param>
	/// <returns></returns>
	[SuppressMessage("Minor Code Smell", "S3872:Parameter names should not duplicate the names of their methods", Justification = "Public API: renaming the parameter would break callers using named arguments.")]
	public ExcelRow Row(int row)
	{
		CheckSheetType();
		return row is < 1 or > ExcelPackage.MaxRows ? throw new ArgumentException("Row number out of bounds") : new ExcelRow(this, row);
	}
	/// <summary>
	/// Provides access to an individual column within the worksheet so you can set its properties.
	/// </summary>
	/// <param name="col">The column number in the worksheet</param>
	/// <returns></returns>
	public ExcelColumn Column(int col)
	{
		CheckSheetType();
		if (col is < 1 or > ExcelPackage.MaxColumns)
		{
			throw new ArgumentException("Column number out of bounds");
		}

		var column = GetValueInner(0, col) as ExcelColumn;
		if (column != null)
		{
			if (column.ColumnMin != column.ColumnMax)
			{
				var maxCol = column.ColumnMax;
				column.ColumnMax = col;
				CopyColumn(column, col + 1, maxCol);
			}
		}
		else
		{
			int r = 0, c = col;
			if (_values.PrevCell(ref r, ref c))
			{
				column = GetValueInner(0, c) as ExcelColumn;
				var maxCol = column.ColumnMax;
				if (maxCol >= col)
				{
					column.ColumnMax = col - 1;
					if (maxCol > col)
					{
						CopyColumn(column, col + 1, maxCol);
					}

					return CopyColumn(column, col, col);
				}
			}

			column = new ExcelColumn(this, col);
			SetValueInner(0, col, column);
		}

		return column;
	}

	/// <summary>
	/// Returns the name of the worksheet
	/// </summary>
	/// <returns>The name of the worksheet</returns>
	public override string ToString() => Name;
}

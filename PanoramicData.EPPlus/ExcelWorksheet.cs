/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Initial Release		        2011-11-02
 * Jan Källman          Total rewrite               2010-03-01
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table.PivotTable;
using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// Worksheet hidden enumeration
/// </summary>
public enum eWorkSheetHidden
{
	/// <summary>
	/// The worksheet is visible
	/// </summary>
	Visible,
	/// <summary>
	/// The worksheet is hidden but can be shown by the user via the user interface
	/// </summary>
	Hidden,
	/// <summary>
	/// The worksheet is hidden and cannot be shown by the user via the user interface
	/// </summary>
	VeryHidden
}
[Flags]
internal enum CellFlags
{
	RichText = 0x2,
	SharedFormula = 0x4,
	ArrayFormula = 0x8
}
/// <summary>
/// For Cell value structure (for memory optimization of huge sheet)
/// </summary>
[SuppressMessage("Minor Code Smell", "S3898:Value types should implement IEquatable<T>", Justification = "Public API: adding IEquatable to this published struct is out of scope for a behaviour-preserving change.")]
public struct ExcelCoreValue
{
	internal object _value;
	internal int _styleId;
}
/// <summary>
/// Represents an Excel Chartsheet and provides access to its properties and methods
/// </summary>
public class ExcelChartsheet : ExcelWorksheet
{
	public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden, eChartType chartType, ExcelPivotTable pivotTableSource) :
		base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
	{
		Drawings.AddChart("Chart 1", chartType, pivotTableSource);
	}
	public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden) :
		base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
	{
	}
	public ExcelChart Chart => (ExcelChart)Drawings[0];
}
/// <summary>
/// Represents an Excel worksheet and provides access to its properties and methods
/// </summary>
public partial class ExcelWorksheet : XmlHelper, IEqualityComparer<ExcelWorksheet>, IDisposable
{
	internal CellStore<ExcelCoreValue>? _values;
	internal CellStore<object>? _formulas;
	internal FlagCellStore? _flags;
	internal CellStore<List<Token>>? _formulaTokens;

	internal CellStore<Uri>? _hyperLinks;
	internal CellStore<int>? _commentsStore;

	internal Dictionary<int, Formulas>? _sharedFormulas = [];
	internal int _minCol = ExcelPackage.MaxColumns;
	internal int _maxCol;
	internal ExcelPackage? _package;
	private readonly Uri _worksheetUri;
	private string _name;
	private readonly int _sheetID;
	private int _positionID;
	private readonly string _relationshipID;
	private XmlDocument _worksheetXml;
	internal ExcelWorksheetView? _sheetView;
	internal ExcelHeaderFooter _headerFooter;
	/// <summary>
	/// A worksheet
	/// </summary>
	/// <param name="ns">Namespacemanager</param>
	/// <param name="excelPackage">Package</param>
	/// <param name="relID">Relationship ID</param>
	/// <param name="uriWorksheet">URI</param>
	/// <param name="sheetName">Name of the sheet</param>
	/// <param name="sheetID">Sheet id</param>
	/// <param name="positionID">Position</param>
	/// <param name="hide">hide</param>
	public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID,
						  Uri uriWorksheet, string sheetName, int sheetID, int positionID,
						  eWorkSheetHidden hide) :
		base(ns, null)
	{
		SchemaNodeOrder = ["sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "activeXControls", "webPublishItems", "tableParts", "extLst"];
		_package = excelPackage;
		_relationshipID = relID;
		_worksheetUri = uriWorksheet;
		_name = sheetName;
		_sheetID = sheetID;
		_positionID = positionID;
		Hidden = hide;

		/**** Cellstore ****/
		_values = new CellStore<ExcelCoreValue>();
		_formulas = new CellStore<object>();
		_flags = new FlagCellStore();
		_commentsStore = new CellStore<int>();
		_hyperLinks = new CellStore<Uri>();

		_names = new ExcelNamedRangeCollection(Workbook, this);

		CreateXml();
		TopNode = _worksheetXml.DocumentElement;
	}

	/// <summary>
	/// The Uri to the worksheet within the package
	/// </summary>
	internal Uri WorksheetUri => (_worksheetUri);
	/// <summary>
	/// The Zip.ZipPackagePart for the worksheet within the package
	/// </summary>
	internal Packaging.ZipPackagePart Part => (_package.Package.GetPart(WorksheetUri));
	/// <summary>
	/// The ID for the worksheet's relationship with the workbook in the package
	/// </summary>
	internal string RelationshipID => (_relationshipID);
	/// <summary>
	/// The unique identifier for the worksheet.
	/// </summary>
	internal int SheetID => (_sheetID);
	/// <summary>
	/// The position of the worksheet.
	/// </summary>
	internal int PositionID { get { return _positionID; } set { _positionID = value; } }
	/// <summary>
	/// The index in the worksheets collection
	/// </summary>
	public int Index => (_positionID);
	/// <summary>
	/// The workbook object
	/// </summary>
	public ExcelWorkbook Workbook => _package.Workbook;

	private static void DisposeInternal(IDisposable candidateDisposable) => candidateDisposable?.Dispose();


	public void Dispose()
	{
		DisposeInternal(_values);
		DisposeInternal(_formulas);
		DisposeInternal(_flags);
		DisposeInternal(_hyperLinks);
		DisposeInternal(_commentsStore);
		DisposeInternal(_formulaTokens);

		_values = null;
		_formulas = null;
		_flags = null;
		_hyperLinks = null;
		_commentsStore = null;
		_formulaTokens = null;

		_package = null;
		_pivotTables = null;
		_protection = null;
		_sharedFormulas?.Clear();
		_sharedFormulas = null;
		_sheetView = null;
		_tables = null;
		_vmlDrawings = null;
		_conditionalFormatting = null;
		_dataValidation = null;
		_drawings = null;
	}

	public bool Equals(ExcelWorksheet x, ExcelWorksheet y) => x.Name == y.Name && x.SheetID == y.SheetID && x.WorksheetXml.OuterXml == y.WorksheetXml.OuterXml;

	public int GetHashCode(ExcelWorksheet obj) => obj.WorksheetXml.OuterXml.GetHashCode();
}  // END class Worksheet

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
<<<<<<< HEAD
 * Jan K�llman		    Initial Release		       2011-01-01
 * Jan K�llman		    License changed GPL-->LGPL 2011-12-27
 * Richard Tallent		Fix escaping of quotes     2012-10-31
=======
 * Jan Källman		    Initial Release		       2011-01-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 * Richard Tallent		Fix escaping of quotes					2012-10-31
>>>>>>> 21c1f80b2e27bc423e3de7f9f2e2b8c9d63934f2
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System.Xml;

namespace OfficeOpenXml;

#region Public Enum ExcelCalcMode
/// <summary>
/// How the application should calculate formulas in the workbook
/// </summary>
public enum ExcelCalcMode
{
	/// <summary>
	/// Indicates that calculations in the workbook are performed automatically when cell values change. 
	/// The application recalculates those cells that are dependent on other cells that contain changed values. 
	/// This mode of calculation helps to avoid unnecessary calculations.
	/// </summary>
	Automatic,
	/// <summary>
	/// Indicates tables be excluded during automatic calculation
	/// </summary>
	AutomaticNoTable,
	/// <summary>
	/// Indicates that calculations in the workbook be triggered manually by the user. 
	/// </summary>
	Manual
}
#endregion

/// <summary>
/// Represents the Excel workbook and provides access to all the 
/// document properties and worksheets within the workbook.
/// </summary>
public sealed partial class ExcelWorkbook : XmlHelper, IDisposable
{
	internal class SharedStringItem
	{
		internal int pos;
		internal string Text;
		internal bool isRichText;
	}
	internal ExcelPackage _package;
	internal ExcelWorksheets _worksheets;
	private OfficeProperties _properties;

	private ExcelStyles _styles;

	/// <summary>
	/// Creates a new instance of the ExcelWorkbook class.
	/// </summary>
	/// <param name="package">The parent package</param>
	/// <param name="namespaceManager">NamespaceManager</param>
	internal ExcelWorkbook(ExcelPackage package, XmlNamespaceManager namespaceManager) :
		base(namespaceManager)
	{
		_package = package;
		WorkbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
		SharedStringsUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
		StylesUri = new Uri("/xl/styles.xml", UriKind.Relative);

		_names = new ExcelNamedRangeCollection(this);
		_namespaceManager = namespaceManager;
		TopNode = WorkbookXml.DocumentElement;
		SchemaNodeOrder = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "functionPrototypes", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
		FullCalcOnLoad = true;  //Full calculation on load by default, for both new workbooks and templates.
		GetSharedStrings();
	}

	internal Dictionary<string, SharedStringItem> _sharedStrings = []; //Used when reading cells.
	internal List<SharedStringItem> _sharedStringsList = []; //Used when reading cells.
	internal ExcelNamedRangeCollection _names;
	internal int _nextDrawingID;
	internal int _nextTableID = int.MinValue;
	internal int _nextPivotTableID = int.MinValue;
	internal XmlNamespaceManager _namespaceManager;
	internal FormulaParser _formulaParser;
	internal FormulaParserManager _parserManager;
	internal CellStore<List<Token>> _formulaTokens;
	/// <summary>
	/// Provides access to all the worksheets in the workbook.
	/// Note: Worksheets index either starts by 0 or 1 depending on the Excelpackage.Compatibility.IsWorksheets1Based property.
	/// Default is 1 for .Net 3.5 and .Net 4 and 0 for .Net Core.
	/// </summary>
	public ExcelWorksheets Worksheets
	{
		get
		{
			if (_worksheets == null)
			{
				var sheetsNode = _workbookXml.DocumentElement.SelectSingleNode("d:sheets", _namespaceManager);
				sheetsNode ??= CreateNode("d:sheets");

				_worksheets = new ExcelWorksheets(_package, _namespaceManager, sheetsNode);
			}

			return _worksheets;
		}
	}

	/// <summary>
	/// Provides access to named ranges
	/// </summary>
	public ExcelNamedRangeCollection Names => _names;
	decimal _standardFontWidth = decimal.MinValue;
	string _fontID = "";
	internal FormulaParser FormulaParser
	{
		get
		{
			_formulaParser ??= new FormulaParser(new EpplusExcelDataProvider(_package));

			return _formulaParser;
		}
	}

	public FormulaParserManager FormulaParserManager
	{
		get
		{
			_parserManager ??= new FormulaParserManager(FormulaParser);

			return _parserManager;
		}
	}

	ExcelProtection _protection;
	/// <summary>
	/// Access properties to protect or unprotect a workbook
	/// </summary>
	public ExcelProtection Protection
	{
		get
		{
			_protection ??= new ExcelProtection(NameSpaceManager, TopNode, this)
			{
				SchemaNodeOrder = SchemaNodeOrder
			};

			return _protection;
		}
	}
	ExcelWorkbookView _view;
	/// <summary>
	/// Access to workbook view properties
	/// </summary>
	public ExcelWorkbookView View
	{
		get
		{
			_view ??= new ExcelWorkbookView(NameSpaceManager, TopNode, this);

			return _view;
		}
	}
	ExcelVbaProject _vba;
	/// <summary>
	/// URI to the workbook inside the package
	/// </summary>
	internal Uri WorkbookUri { get; private set; }
	/// <summary>
	/// URI to the styles inside the package
	/// </summary>
	internal Uri StylesUri { get; private set; }
	/// <summary>
	/// URI to the shared strings inside the package
	/// </summary>
	internal Uri SharedStringsUri { get; private set; }
	/// <summary>
	/// Returns a reference to the workbook's part within the package
	/// </summary>
	internal Packaging.ZipPackagePart Part => (_package.Package.GetPart(WorkbookUri));

	private XmlDocument _workbookXml;
	/// <summary>
	/// Provides access to the XML data representing the workbook in the package.
	/// </summary>
	public XmlDocument WorkbookXml
	{
		get
		{
			if (_workbookXml == null)
			{
				CreateWorkbookXml(_namespaceManager);
			}

			return _workbookXml;
		}
	}
	const string codeModuleNamePath = "d:workbookPr/@codeName";
	internal string CodeModuleName
	{
		get
		{
			return GetXmlNodeString(codeModuleNamePath);
		}
		set
		{
			SetXmlNodeString(codeModuleNamePath, value);
		}
	}
	internal void CodeNameChange(string value) => CodeModuleName = value;
	public VBA.ExcelVBAModule CodeModule => VbaProject != null ? VbaProject.Modules[CodeModuleName] : null;

	const string date1904Path = "d:workbookPr/@date1904";
	internal const double date1904Offset = 365.5 * 4;  // offset to fix 1900 and 1904 differences, 4 OLE years
	private bool? date1904Cache;
	private XmlDocument _stylesXml;
	/// <summary>
	/// Package styles collection. Used internally to access style data.
	/// </summary>
	public ExcelStyles Styles
	{
		get
		{
			_styles ??= new ExcelStyles(NameSpaceManager, StylesXml, this);

			return _styles;
		}
	}

	/// <summary>
	/// The office document properties
	/// </summary>
	public OfficeProperties Properties
	{
		get
		{
			//  Create a NamespaceManager to handle the default namespace, 
			//  and create a prefix for the default namespace:                   
			_properties ??= new OfficeProperties(_package, NameSpaceManager);

			return _properties;
		}
	}

	private readonly string CALC_MODE_PATH = "d:calcPr/@calcMode";

	private const string FULL_CALC_ON_LOAD_PATH = "d:calcPr/@fullCalcOnLoad";
	internal void AddPivotTable(string cacheID, Uri defUri)
	{
		CreateNode("d:pivotCaches");

		var item = WorkbookXml.CreateElement("pivotCache", ExcelPackage.schemaMain);
		item.SetAttribute("cacheId", cacheID);
		var rel = Part.CreateRelationship(UriHelper.ResolvePartUri(WorkbookUri, defUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
		item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

		var pivotCaches = WorkbookXml.SelectSingleNode("//d:pivotCaches", NameSpaceManager);
		pivotCaches.AppendChild(item);
	}
	internal List<string> _externalReferences = [];

	public void Dispose()
	{
		if (_sharedStrings != null)
		{
			_sharedStrings.Clear();
			_sharedStrings = null;
		}

		if (_sharedStringsList != null)
		{
			_sharedStringsList.Clear();
			_sharedStringsList = null;
		}

		_vba = null;
		if (_worksheets != null)
		{
			_worksheets.Dispose();
			_worksheets = null;
		}

		_package = null;
		_properties = null;
		if (_formulaParser != null)
		{
			_formulaParser.Dispose();
			_formulaParser = null;
		}
	}
} // end Workbook

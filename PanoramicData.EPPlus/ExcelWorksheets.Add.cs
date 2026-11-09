using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.VBA;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	/// <summary>
	/// Adds a new blank worksheet.
	/// </summary>
	/// <param name="Name">The name of the workbook</param>
	public ExcelWorksheet Add(string Name)
	{
		var worksheet = AddSheet(Name, false, null);
		return worksheet;
	}
	private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType, ExcelPivotTable pivotTableSource = null)
	{
		lock (_worksheets)
		{
			Name = ValidateFixSheetName(Name);
			if (GetByName(Name) != null)
			{
				throw new InvalidOperationException(ERR_DUP_WORKSHEET + " : " + Name);
			}

			GetSheetURI(ref Name, out var sheetID, out var uriWorksheet, isChart);
			var worksheetPart = _pck.Package.CreatePart(uriWorksheet, isChart ? CHARTSHEET_CONTENTTYPE : WORKSHEET_CONTENTTYPE, _pck.Compression);

			//Create the new, empty worksheet and save it to the package
			StreamWriter streamWorksheet = new(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
			var worksheetXml = CreateNewWorksheet(isChart);
			worksheetXml.Save(streamWorksheet);
			Packaging.ZipPackage.Flush();

			var rel = CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart);

			var positionID = _worksheets.Count + _pck._worksheetAdd;
			ExcelWorksheet worksheet;
			if (isChart)
			{
				worksheet = new ExcelChartsheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible, (eChartType)chartType, pivotTableSource);
			}
			else
			{
				worksheet = new ExcelWorksheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);
			}

			_worksheets.Add(positionID, worksheet);
			if (_pck.Workbook.VbaProject != null)
			{
				var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(worksheet);
				_pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(worksheet.CodeNameChange) { Name = name, Code = "", Attributes = ExcelVbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
				worksheet.CodeModuleName = name;

			}

			return worksheet;
		}
	}
	/// <summary>
	/// Adds a copy of a worksheet
	/// </summary>
	/// <param name="Name">The name of the workbook</param>
	/// <param name="Copy">The worksheet to be copied</param>
	public ExcelWorksheet Add(string Name, ExcelWorksheet Copy)
	{
		lock (_worksheets)
		{
			if (Copy is ExcelChartsheet)
			{
				throw new ArgumentException("Can not copy a chartsheet");
			}

			if (GetByName(Name) != null)
			{
				throw new InvalidOperationException(ERR_DUP_WORKSHEET);
			}

			GetSheetURI(ref Name, out var sheetID, out var uriWorksheet, false);

			//Create a copy of the worksheet XML
			var worksheetPart = _pck.Package.CreatePart(uriWorksheet, WORKSHEET_CONTENTTYPE, _pck.Compression);
			StreamWriter streamWorksheet = new(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
			XmlDocument worksheetXml = new();
			worksheetXml.LoadXml(Copy.WorksheetXml.OuterXml);
			worksheetXml.Save(streamWorksheet);
			Packaging.ZipPackage.Flush();


			//Create a relation to the workbook
			var relID = CreateWorkbookRel(Name, sheetID, uriWorksheet, false);
			ExcelWorksheet added = new(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, _worksheets.Count + _pck._worksheetAdd, eWorkSheetHidden.Visible);

			//Copy comments
			if (Copy.Comments.Count > 0)
			{
				CopyComment(Copy, added);
			}
			else if (Copy.VmlDrawingsComments.Count > 0)    //Vml drawings are copied as part of the comments. 
			{
				CopyVmlDrawing(Copy, added);
			}

			//Copy HeaderFooter
			CopyHeaderFooterPictures(Copy, added);

			if (Copy.Drawings.Count > 0)
			{
				CopyDrawing(Copy, added);
			}

			if (Copy.Tables.Count > 0)
			{
				CopyTable(Copy, added);
			}

			if (Copy.PivotTables.Count > 0)
			{
				CopyPivotTable(Copy, added);
			}

			if (Copy.Names.Count > 0)
			{
				CopySheetNames(Copy, added);
			}

			//Copy all cells
			CloneCells(Copy, added);

			//Copy the VBA code
			if (_pck.Workbook.VbaProject != null)
			{
				var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(added);
				_pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(added.CodeNameChange) { Name = name, Code = Copy.CodeModule.Code, Attributes = ExcelVbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
				Copy.CodeModuleName = name;
			}

			_worksheets.Add(_worksheets.Count + _pck._worksheetAdd, added);

			//Remove any relation to printersettings.
			var pageSetup = added.WorksheetXml.SelectSingleNode("//d:pageSetup", _namespaceManager);
			if (pageSetup != null)
			{
				var attr = (XmlAttribute)pageSetup.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
				if (attr != null)
				{
					// first delete the attribute from the XML
					pageSetup.Attributes.Remove(attr);
				}
			}

			return added;
		}
	}
	/// <summary>
	/// Adds a chartsheet to the workbook.
	/// </summary>
	/// <param name="Name">The name of the worksheet</param>
	/// <param name="chartType">The type of chart</param>
	/// <returns></returns>
	public ExcelChartsheet AddChart(string Name, eChartType chartType) => (ExcelChartsheet)AddSheet(Name, true, chartType, null);
	/// <summary>
	/// Adds a chartsheet to the workbook.
	/// </summary>
	/// <param name="Name">The name of the worksheet</param>
	/// <param name="chartType">The type of chart</param>
	/// <param name="pivotTableSource">The pivottable source</param>
	/// <returns></returns>
	public ExcelChartsheet AddChart(string Name, eChartType chartType, ExcelPivotTable pivotTableSource) => (ExcelChartsheet)AddSheet(Name, true, chartType, pivotTableSource);
	/// <summary>
	/// Copies the named worksheet and creates a new worksheet in the same workbook
	/// </summary>
	/// <param name="Name">The name of the existing worksheet</param>
	/// <param name="NewName">The name of the new worksheet to create</param>
	/// <returns>The new copy added to the end of the worksheets collection</returns>
	public ExcelWorksheet Copy(string Name, string NewName)
	{
		var Copy = this[Name];
		if (Copy == null)
		{
			throw new ArgumentException(string.Format("Copy worksheet error: Could not find worksheet to copy '{0}'", Name));
		}

		var added = Add(NewName, Copy);
		return added;
	}
}

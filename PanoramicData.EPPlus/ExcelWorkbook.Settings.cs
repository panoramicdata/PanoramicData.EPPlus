using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System.Text;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class ExcelWorkbook
{
	/// <summary>
	/// A reference to the VBA project.
	/// Null if no project exists.
	/// Use Workbook.CreateVBAProject to create a new VBA-Project
	/// </summary>
	public ExcelVbaProject VbaProject
	{
		get
		{
			if (_vba == null)
			{
				if (_package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
				{
					_vba = new ExcelVbaProject(this);
				}
			}

			return _vba;
		}
	}

	/// <summary>
	/// Create an empty VBA project.
	/// </summary>
	public void CreateVBAProject()
	{
		if (_vba != null || _package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
		{
			throw new InvalidOperationException("VBA project already exists.");
		}

		_vba = new ExcelVbaProject(this);
		_vba.Create();
	}

	/// <summary>
	/// The date systems used by Microsoft Excel can be based on one of two different dates. By default, a serial number of 1 in Microsoft Excel represents January 1, 1900.
	/// The default for the serial number 1 can be changed to represent January 2, 1904.
	/// This option was included in Microsoft Excel for Windows to make it compatible with Excel for the Macintosh, which defaults to January 2, 1904.
	/// </summary>
	public bool Date1904
	{
		get
		{
			date1904Cache ??= GetXmlNodeBool(date1904Path, false);

			return date1904Cache.Value;
		}
		set
		{
			if (Date1904 != value)
			{
				// Like Excel when the option it's changed update it all cells with Date format
				foreach (var item in Worksheets)
				{
					item.UpdateCellsWithDate1904Setting();
				}
			}

			date1904Cache = value;
			SetXmlNodeBool(date1904Path, value, false);
		}
	}


	/// <summary>
	/// Create or read the XML for the workbook.
	/// </summary>
	private void CreateWorkbookXml(XmlNamespaceManager namespaceManager)
	{
		if (_package.Package.PartExists(WorkbookUri))
		{
			_workbookXml = _package.GetXmlFromUri(WorkbookUri);
		}
		else
		{
			// create a new workbook part and add to the package
			var partWorkbook = _package.Package.CreatePart(WorkbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", _package.Compression);

			// create the workbook
			_workbookXml = new XmlDocument(namespaceManager.NameTable)
			{
				PreserveWhitespace = ExcelPackage.preserveWhitespace
			};
			// create the workbook element
			var wbElem = _workbookXml.CreateElement("workbook", ExcelPackage.schemaMain);

			// Add the relationships namespace
			wbElem.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);

			_workbookXml.AppendChild(wbElem);

			// create the bookViews and workbooks element
			var bookViews = _workbookXml.CreateElement("bookViews", ExcelPackage.schemaMain);
			wbElem.AppendChild(bookViews);
			var workbookView = _workbookXml.CreateElement("workbookView", ExcelPackage.schemaMain);
			bookViews.AppendChild(workbookView);

			// save it to the package
			StreamWriter stream = new(partWorkbook.GetStream(FileMode.Create, FileAccess.Write));
			_workbookXml.Save(stream);
			Packaging.ZipPackage.Flush();
		}
	}
	/// <summary>
	/// Provides access to the XML data representing the styles in the package. 
	/// </summary>
	public XmlDocument StylesXml
	{
		get
		{
			if (_stylesXml == null)
			{
				if (_package.Package.PartExists(StylesUri))
				{
					_stylesXml = _package.GetXmlFromUri(StylesUri);
				}
				else
				{
					// create a new styles part and add to the package
					var part = _package.Package.CreatePart(StylesUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", _package.Compression);
					// create the style sheet

					StringBuilder xml = new("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
					xml.Append("<numFmts />");
					xml.Append("<fonts count=\"1\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts>");
					xml.Append("<fills><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills>");
					xml.Append("<borders><border><left /><right /><top /><bottom /><diagonal /></border></borders>");
					xml.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs>");
					xml.Append("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" xfId=\"0\" /></cellXfs>");
					xml.Append("<cellStyles><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles>");
					xml.Append("<dxfs count=\"0\" />");
					xml.Append("</styleSheet>");

					_stylesXml = new XmlDocument();
					_stylesXml.LoadXml(xml.ToString());

					//Save it to the package
					StreamWriter stream = new(part.GetStream(FileMode.Create, FileAccess.Write));

					_stylesXml.Save(stream);
					Packaging.ZipPackage.Flush();

					// create the relationship between the workbook and the new shared strings part
					_package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, StylesUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/styles");
					Packaging.ZipPackage.Flush();
				}
			}

			return _stylesXml;
		}
		set
		{
			_stylesXml = value;
		}
	}
	/// <summary>
	/// Calculation mode for the workbook.
	/// </summary>
	public ExcelCalcMode CalcMode
	{
		get
		{
			var calcMode = GetXmlNodeString(CALC_MODE_PATH);
			return calcMode switch
			{
				"autoNoTable" => ExcelCalcMode.AutomaticNoTable,
				"manual" => ExcelCalcMode.Manual,
				_ => ExcelCalcMode.Automatic,
			};
		}
		set
		{
			switch (value)
			{
				case ExcelCalcMode.AutomaticNoTable:
					SetXmlNodeString(CALC_MODE_PATH, "autoNoTable");
					break;
				case ExcelCalcMode.Manual:
					SetXmlNodeString(CALC_MODE_PATH, "manual");
					break;
				default:
					SetXmlNodeString(CALC_MODE_PATH, "auto");
					break;

			}
		}
	}
	/// <summary>
	/// Should Excel do a full calculation after the workbook has been loaded?
	/// <remarks>This property is always true for both new workbooks and loaded templates(on load). If this is not the wanted behavior set this property to false.</remarks>
	/// </summary>
	public bool FullCalcOnLoad
	{
		get
		{
			return GetXmlNodeBool(FULL_CALC_ON_LOAD_PATH);
		}
		set
		{
			SetXmlNodeBool(FULL_CALC_ON_LOAD_PATH, value);
		}
	}
}

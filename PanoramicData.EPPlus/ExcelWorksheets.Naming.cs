using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet, bool isChart)
	{
		//Create the relationship between the workbook and the new worksheet
		var rel = _pck.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_pck.Workbook.WorkbookUri, uriWorksheet), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/" + (isChart ? "chartsheet" : "worksheet"));
		Packaging.ZipPackage.Flush();

		//Create the new sheet node
		var worksheetNode = _pck.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
		worksheetNode.SetAttribute("name", Name);
		worksheetNode.SetAttribute("sheetId", sheetID.ToString());
		worksheetNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

		TopNode.AppendChild(worksheetNode);
		return rel.Id;
	}

	private void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet, bool isChart)
	{
		Name = ValidateFixSheetName(Name);
		sheetID = this.Any() ? this.Max(ws => ws.SheetID) + 1 : 1;
		var uriId = sheetID;


		// get the next available worhsheet uri
		do
		{
			if (isChart)
			{
				uriWorksheet = new Uri("/xl/chartsheets/chartsheet" + uriId + ".xml", UriKind.Relative);
			}
			else
			{
				uriWorksheet = new Uri("/xl/worksheets/sheet" + uriId + ".xml", UriKind.Relative);
			}

			uriId++;
		} while (_pck.Package.PartExists(uriWorksheet));
	}

	internal static string ValidateFixSheetName(string Name)
	{
		//remove invalid characters
		if (ValidateName(Name))
		{
			if (Name.IndexOf(':') > -1)
			{
				Name = Name.Replace(":", " ");
			}

			if (Name.IndexOf('/') > -1)
			{
				Name = Name.Replace("/", " ");
			}

			if (Name.IndexOf('\\') > -1)
			{
				Name = Name.Replace("\\", " ");
			}

			if (Name.IndexOf('?') > -1)
			{
				Name = Name.Replace("?", " ");
			}

			if (Name.IndexOf('[') > -1)
			{
				Name = Name.Replace("[", " ");
			}

			if (Name.IndexOf(']') > -1)
			{
				Name = Name.Replace("]", " ");
			}
		}

		if (Name.Trim() == "")
		{
			throw new ArgumentException("The worksheet can not have an empty name");
		}

		if (Name.StartsWith('\'') || Name.EndsWith('\''))
		{
			throw new ArgumentException("The worksheet name can not start or end with an apostrophe.");
		}

		if (Name.Length > 31)
		{
			Name = Name[..31];   //A sheet can have max 31 char's            
		}

		return Name;
	}
	/// <summary>
	/// Validate the sheetname
	/// </summary>
	/// <param name="Name">The Name</param>
	/// <returns>True if valid</returns>
	private static bool ValidateName(string Name) => System.Text.RegularExpressions.Regex.IsMatch(Name, @":|\?|/|\\|\[|\]");

	/// <summary>
	/// Creates the XML document representing a new empty worksheet
	/// </summary>
	/// <returns></returns>
	internal static XmlDocument CreateNewWorksheet(bool isChart)
	{
		XmlDocument xmlDoc = new();
		var elemWs = xmlDoc.CreateElement(isChart ? "chartsheet" : "worksheet", ExcelPackage.schemaMain);
		elemWs.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
		xmlDoc.AppendChild(elemWs);


		if (isChart)
		{
			var elemSheetPr = xmlDoc.CreateElement("sheetPr", ExcelPackage.schemaMain);
			elemWs.AppendChild(elemSheetPr);

			var elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
			elemWs.AppendChild(elemSheetViews);

			var elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
			elemSheetView.SetAttribute("workbookViewId", "0");
			elemSheetView.SetAttribute("zoomToFit", "1");

			elemSheetViews.AppendChild(elemSheetView);
		}
		else
		{
			var elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
			elemWs.AppendChild(elemSheetViews);

			var elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
			elemSheetView.SetAttribute("workbookViewId", "0");
			elemSheetViews.AppendChild(elemSheetView);

			var elemSheetFormatPr = xmlDoc.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
			elemSheetFormatPr.SetAttribute("defaultRowHeight", "15");
			elemWs.AppendChild(elemSheetFormatPr);

			var elemSheetData = xmlDoc.CreateElement("sheetData", ExcelPackage.schemaMain);
			elemWs.AppendChild(elemSheetData);
		}

		return xmlDoc;
	}
}

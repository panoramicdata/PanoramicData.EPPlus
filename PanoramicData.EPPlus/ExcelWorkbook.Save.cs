using OfficeOpenXml.Packaging.DotNetZip;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace OfficeOpenXml;

public sealed partial class ExcelWorkbook
{
	/// <summary>
	/// Saves the workbook and all its components to the package.
	/// For internal use only!
	/// </summary>
	internal void Save()  // Workbook Save
	{
		if (Worksheets.Count == 0)
		{
			throw new InvalidOperationException("The workbook must contain at least one worksheet");
		}

		DeleteCalcChain();
		UpdateWorkbookContentType();
		UpdateDefinedNamesXml();

		// save the workbook
		if (_workbookXml != null)
		{
			_package.SavePart(WorkbookUri, _workbookXml);
		}

		// save the properties of the workbook
		_properties?.Save();

		// save the style sheet
		Styles.UpdateXml();
		_package.SavePart(StylesUri, _stylesXml);

		SaveWorksheets();

		// Issue 15252: save SharedStrings only once
		GetOrCreateSharedStringsPart().SaveHandler = SaveSharedStringHandler;

		// Data validation
		ValidateDataValidations();

		//VBA
		if (_vba != null)
		{
			VbaProject.Save();
		}

	}
	[SuppressMessage("Major Code Smell", "S3440:Variables should not be checked against the values they're about to be assigned", Justification = "The ContentType setter rewrites the package content-type map, so it is only called when the value changes.")]
	private void UpdateWorkbookContentType()
	{
		if (_vba == null && !_package.Package.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
		{
			if (Part.ContentType != ExcelPackage.contentTypeWorkbookDefault)
			{
				Part.ContentType = ExcelPackage.contentTypeWorkbookDefault;
			}
		}
		else
		{
			if (Part.ContentType != ExcelPackage.contentTypeWorkbookMacroEnabled)
			{
				Part.ContentType = ExcelPackage.contentTypeWorkbookMacroEnabled;
			}
		}
	}

	private void SaveWorksheets()
	{
		// save all the open worksheets
		var isProtected = Protection.LockWindows || Protection.LockStructure;
		foreach (var worksheet in Worksheets)
		{
			if (isProtected && Protection.LockWindows)
			{
				worksheet.View.WindowProtection = true;
			}

			worksheet.Save();
			worksheet.Part.SaveHandler = worksheet.SaveHandler;
		}
	}

	private Packaging.ZipPackagePart GetOrCreateSharedStringsPart()
	{
		if (_package.Package.PartExists(SharedStringsUri))
		{
			return _package.Package.GetPart(SharedStringsUri);
		}

		var part = _package.Package.CreatePart(SharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", _package.Compression);
		Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, SharedStringsUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/sharedStrings");
		return part;
	}

	private void DeleteCalcChain()
	{
		//Remove the calc chain if it exists.
		Uri uriCalcChain = new("/xl/calcChain.xml", UriKind.Relative);
		if (_package.Package.PartExists(uriCalcChain))
		{
			Uri calcChain = new("calcChain.xml", UriKind.Relative);
			foreach (var relationship in _package.Workbook.Part.GetRelationships())
			{
				if (relationship.TargetUri == calcChain)
				{
					_package.Workbook.Part.DeleteRelationship(relationship.Id);
					break;
				}
			}
			// delete the calcChain part
			_package.Package.DeletePart(uriCalcChain);
		}
	}

	private void ValidateDataValidations()
	{
		foreach (var sheet in _package.Workbook.Worksheets)
		{
			if (sheet is not ExcelChartsheet)
			{
				sheet.DataValidations.ValidateAll();
			}
		}
	}

	private void SaveSharedStringHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
	{
		stream.CompressionLevel = (Packaging.DotNetZip.Zlib.CompressionLevel)compressionLevel;
		stream.PutNextEntry(fileName);

		var cache = new StringBuilder();
		var sw = new StreamWriter(stream);
		cache.AppendFormat("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", _sharedStrings.Count);
		foreach (var t in _sharedStrings.Keys)
		{

			AppendSharedString(cache, t, _sharedStrings[t]);

			if (cache.Length > 0x600000)
			{
				sw.Write(cache.ToString());
				cache = new StringBuilder();
			}
		}

		cache.Append("</sst>");
		sw.Write(cache.ToString());
		sw.Flush();
	}

	private static void AppendSharedString(StringBuilder cache, string t, SharedStringItem ssi)
	{
		if (ssi.isRichText)
		{
			cache.Append("<si>");
			ConvertUtil.ExcelEncodeString(cache, t);
			cache.Append("</si>");
		}
		else
		{
			if (t.Length > 0 && (t[0] == ' ' || t[^1] == ' ' || t.Contains("  ") || t.Contains('\t') || t.Contains('\n') || t.Contains('\n')))   //Fixes issue 14849
			{
				cache.Append("<si><t xml:space=\"preserve\">");
			}
			else
			{
				cache.Append("<si><t>");
			}

			ConvertUtil.ExcelEncodeString(cache, ConvertUtil.ExcelEscapeString(t));
			cache.Append("</t></si>");
		}
	}
}

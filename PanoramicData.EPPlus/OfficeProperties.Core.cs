using OfficeOpenXml.Utils;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class OfficeProperties
{
	/// <summary>
	/// Provides access to the XML document that holds all the code 
	/// document properties.
	/// </summary>
	public XmlDocument CorePropertiesXml
	{
		get
		{
			if (_xmlPropertiesCore == null)
			{
				var xml = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><cp:coreProperties xmlns:cp=\"{0}\" xmlns:dc=\"{1}\" xmlns:dcterms=\"{2}\" xmlns:dcmitype=\"{3}\" xmlns:xsi=\"{4}\"></cp:coreProperties>",
					ExcelPackage.schemaCore,
					ExcelPackage.schemaDc,
					ExcelPackage.schemaDcTerms,
					ExcelPackage.schemaDcmiType,
					ExcelPackage.schemaXsi);

				_xmlPropertiesCore = GetXmlDocument(xml, _uriPropertiesCore, @"application/vnd.openxmlformats-package.core-properties+xml", @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
			}

			return _xmlPropertiesCore;
		}
	}

	private XmlDocument GetXmlDocument(string startXml, Uri uri, string contentType, string relationship)
	{
		XmlDocument xmlDoc;
		if (_package.Package.PartExists(uri))
		{
			xmlDoc = _package.GetXmlFromUri(uri);
		}
		else
		{
			xmlDoc = new XmlDocument();
			xmlDoc.LoadXml(startXml);

			// Create a the part and add to the package
			var part = _package.Package.CreatePart(uri, contentType);

			// Save it to the package
			StreamWriter stream = new(part.GetStream(FileMode.Create, FileAccess.Write));
			xmlDoc.Save(stream);
			Packaging.ZipPackage.Flush();

			// create the relationship between the workbook and the new shared strings part
			_package.Package.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), uri), Packaging.TargetMode.Internal, relationship);
			Packaging.ZipPackage.Flush();
		}

		return xmlDoc;
	}
	/// <summary>
	/// Gets/sets the title property of the document (core property)
	/// </summary>
	public string Title
	{
		get { return _coreHelper.GetXmlNodeString(TitlePath); }
		set { _coreHelper.SetXmlNodeString(TitlePath, value); }
	}
	/// <summary>
	/// Gets/sets the subject property of the document (core property)
	/// </summary>
	public string Subject
	{
		get { return _coreHelper.GetXmlNodeString(SubjectPath); }
		set { _coreHelper.SetXmlNodeString(SubjectPath, value); }
	}
	/// <summary>
	/// Gets/sets the author property of the document (core property)
	/// </summary>
	public string Author
	{
		get { return _coreHelper.GetXmlNodeString(AuthorPath); }
		set { _coreHelper.SetXmlNodeString(AuthorPath, value); }
	}
	/// <summary>
	/// Gets/sets the comments property of the document (core property)
	/// </summary>
	public string Comments
	{
		get { return _coreHelper.GetXmlNodeString(CommentsPath); }
		set { _coreHelper.SetXmlNodeString(CommentsPath, value); }
	}
	/// <summary>
	/// Gets/sets the keywords property of the document (core property)
	/// </summary>
	public string Keywords
	{
		get { return _coreHelper.GetXmlNodeString(KeywordsPath); }
		set { _coreHelper.SetXmlNodeString(KeywordsPath, value); }
	}
	/// <summary>
	/// Gets/sets the lastModifiedBy property of the document (core property)
	/// </summary>
	public string LastModifiedBy
	{
		get { return _coreHelper.GetXmlNodeString(LastModifiedByPath); }
		set
		{
			_coreHelper.SetXmlNodeString(LastModifiedByPath, value);
		}
	}
	/// <summary>
	/// Gets/sets the lastPrinted property of the document (core property)
	/// </summary>
	public string LastPrinted
	{
		get { return _coreHelper.GetXmlNodeString(LastPrintedPath); }
		set { _coreHelper.SetXmlNodeString(LastPrintedPath, value); }
	}

	/// <summary>
	/// Gets/sets the created property of the document (core property)
	/// </summary>
	public DateTime Created
	{
		get
		{
			return DateTime.TryParse(_coreHelper.GetXmlNodeString(CreatedPath), out var date) ? date : DateTime.MinValue;
		}
		set
		{
			var dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
			_coreHelper.SetXmlNodeString(CreatedPath, dateString);
			_coreHelper.SetXmlNodeString(CreatedPath + "/@xsi:type", "dcterms:W3CDTF");
		}
	}
	/// <summary>
	/// Gets/sets the category property of the document (core property)
	/// </summary>
	public string Category
	{
		get { return _coreHelper.GetXmlNodeString(CategoryPath); }
		set { _coreHelper.SetXmlNodeString(CategoryPath, value); }
	}
	/// <summary>
	/// Gets/sets the status property of the document (core property)
	/// </summary>
	public string Status
	{
		get { return _coreHelper.GetXmlNodeString(ContentStatusPath); }
		set { _coreHelper.SetXmlNodeString(ContentStatusPath, value); }
	}
}

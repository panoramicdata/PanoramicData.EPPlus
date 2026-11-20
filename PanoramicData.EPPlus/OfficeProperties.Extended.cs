using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class OfficeProperties
{
	/// <summary>
	/// Provides access to the XML document that holds the extended properties of the document (app.xml)
	/// </summary>
	public XmlDocument ExtendedPropertiesXml
	{
		get
		{
			_xmlPropertiesExtended ??= GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
						ExcelPackage.schemaVt,
						ExcelPackage.schemaExtended),
					_uriPropertiesExtended,
					@"application/vnd.openxmlformats-officedocument.extended-properties+xml",
					@"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");

			return _xmlPropertiesExtended;
		}
	}
	/// <summary>
	/// Gets/Set the Application property of the document (extended property)
	/// </summary>
	public string Application
	{
		get { return _extendedHelper.GetXmlNodeString(ApplicationPath); }
		set { _extendedHelper.SetXmlNodeString(ApplicationPath, value); }
	}
	/// <summary>
	/// Gets/sets the HyperlinkBase property of the document (extended property)
	/// </summary>
	public Uri HyperlinkBase
	{
		get { return new Uri(_extendedHelper.GetXmlNodeString(HyperlinkBasePath), UriKind.Absolute); }
		set { _extendedHelper.SetXmlNodeString(HyperlinkBasePath, value.AbsoluteUri); }
	}
	/// <summary>
	/// Gets/Set the AppVersion property of the document (extended property)
	/// </summary>
	public string AppVersion
	{
		get { return _extendedHelper.GetXmlNodeString(AppVersionPath); }
		set { _extendedHelper.SetXmlNodeString(AppVersionPath, value); }
	}

	/// <summary>
	/// Gets/sets the Company property of the document (extended property)
	/// </summary>
	public string Company
	{
		get { return _extendedHelper.GetXmlNodeString(CompanyPath); }
		set { _extendedHelper.SetXmlNodeString(CompanyPath, value); }
	}
	/// <summary>
	/// Gets/sets the Manager property of the document (extended property)
	/// </summary>
	public string Manager
	{
		get { return _extendedHelper.GetXmlNodeString(ManagerPath); }
		set { _extendedHelper.SetXmlNodeString(ManagerPath, value); }
	}
	/// <summary>
	/// Gets/sets the modified property of the document (core property)
	/// </summary>
	public DateTime Modified
	{
		get
		{
			return DateTime.TryParse(_coreHelper.GetXmlNodeString(ModifiedPath), out var date) ? date : DateTime.MinValue;
		}
		set
		{
			var dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
			_coreHelper.SetXmlNodeString(ModifiedPath, dateString);
			_coreHelper.SetXmlNodeString(ModifiedPath + "/@xsi:type", "dcterms:W3CDTF");
		}
	}
	/// <summary>
	/// Indicates whether hyperlinks in a document are up-to-date
	/// </summary>
	public bool LinksUpToDate
	{
		get { return _extendedHelper.GetXmlNodeBool(LinksUpToDatePath); }
		set { _extendedHelper.SetXmlNodeBool(LinksUpToDatePath, value); }
	}
	/// <summary>
	/// Hyperlinks need update
	/// </summary>
	public bool HyperlinksChanged
	{
		get { return _extendedHelper.GetXmlNodeBool(HyperlinksChangedPath); }
		set { _extendedHelper.SetXmlNodeBool(HyperlinksChangedPath, value); }
	}
	/// <summary>
	/// Display mode of the document thumbnail. True to enable scaling. False to enable cropping.
	/// </summary>
	public bool ScaleCrop
	{
		get { return _extendedHelper.GetXmlNodeBool(ScaleCropPath); }
		set { _extendedHelper.SetXmlNodeBool(ScaleCropPath, value); }
	}
	/// <summary>
	/// If true, document is shared between multiple producers.
	/// </summary>
	public bool SharedDoc
	{
		get { return _extendedHelper.GetXmlNodeBool(SharedDocPath); }
		set { _extendedHelper.SetXmlNodeBool(SharedDocPath, value); }
	}

	/// <summary>
	/// Get the value of an extended property 
	/// </summary>
	/// <param name="propertyName">The name of the property</param>
	/// <returns>The value</returns>
	public string GetExtendedPropertyValue(string propertyName)
	{
		string retValue = null;
		var searchString = string.Format("xp:Properties/xp:{0}", propertyName);
		var node = ExtendedPropertiesXml.SelectSingleNode(searchString, NameSpaceManager);
		if (node != null)
		{
			retValue = node.InnerText;
		}

		return retValue;
	}
	/// <summary>
	/// Set the value for an extended property
	/// </summary>
	/// <param name="propertyName">The name of the property</param>
	/// <param name="value">The value</param>
	public void SetExtendedPropertyValue(string propertyName, string value)
	{
		var propertyPath = string.Format("xp:Properties/xp:{0}", propertyName);
		_extendedHelper.SetXmlNodeString(propertyPath, value);
	}
}

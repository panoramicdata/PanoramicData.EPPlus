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
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman                      Total rewrite               2010-03-01
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 * Raziq York                       Added Created & Modified    2014-08-20
 *******************************************************************************/
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// Provides access to the properties bag of the package
/// </summary>
public sealed partial class OfficeProperties : XmlHelper
{
	private XmlDocument _xmlPropertiesCore;
	private XmlDocument _xmlPropertiesExtended;
	private XmlDocument _xmlPropertiesCustom;

	private readonly Uri _uriPropertiesCore = new("/docProps/core.xml", UriKind.Relative);
	private readonly Uri _uriPropertiesExtended = new("/docProps/app.xml", UriKind.Relative);
	private readonly Uri _uriPropertiesCustom = new("/docProps/custom.xml", UriKind.Relative);

	readonly XmlHelper _coreHelper;
	readonly XmlHelper _extendedHelper;
	readonly XmlHelper _customHelper;
	private readonly ExcelPackage _package;

	/// <summary>
	/// Provides access to all the office document properties.
	/// </summary>
	/// <param name="package"></param>
	/// <param name="ns"></param>
	internal OfficeProperties(ExcelPackage package, XmlNamespaceManager ns) :
		base(ns)
	{
		_package = package;

		_coreHelper = XmlHelperFactory.Create(ns, CorePropertiesXml.SelectSingleNode("cp:coreProperties", NameSpaceManager));
		_extendedHelper = XmlHelperFactory.Create(ns, ExtendedPropertiesXml);
		_customHelper = XmlHelperFactory.Create(ns, CustomPropertiesXml);

	}
	const string TitlePath = "dc:title";

	const string SubjectPath = "dc:subject";

	const string AuthorPath = "dc:creator";

	const string CommentsPath = "dc:description";

	const string KeywordsPath = "cp:keywords";

	const string LastModifiedByPath = "cp:lastModifiedBy";

	const string LastPrintedPath = "cp:lastPrinted";

	const string CreatedPath = "dcterms:created";

	const string CategoryPath = "cp:category";

	const string ContentStatusPath = "cp:contentStatus";

	const string ApplicationPath = "xp:Properties/xp:Application";

	const string HyperlinkBasePath = "xp:Properties/xp:HyperlinkBase";

	const string AppVersionPath = "xp:Properties/xp:AppVersion";
	const string CompanyPath = "xp:Properties/xp:Company";

	const string ManagerPath = "xp:Properties/xp:Manager";

	const string ModifiedPath = "dcterms:modified";
	const string LinksUpToDatePath = "xp:Properties/xp:LinksUpToDate";
	const string HyperlinksChangedPath = "xp:Properties/xp:HyperlinksChanged";
	const string ScaleCropPath = "xp:Properties/xp:ScaleCrop";

	const string SharedDocPath = "xp:Properties/xp:SharedDoc";

	/// <summary>
	/// Saves the document properties back to the package.
	/// </summary>
	internal void Save()
	{
		if (_xmlPropertiesCore != null)
		{
			_package.SavePart(_uriPropertiesCore, _xmlPropertiesCore);
		}

		if (_xmlPropertiesExtended != null)
		{
			_package.SavePart(_uriPropertiesExtended, _xmlPropertiesExtended);
		}

		if (_xmlPropertiesCustom != null)
		{
			_package.SavePart(_uriPropertiesCustom, _xmlPropertiesCustom);
		}

	}

}

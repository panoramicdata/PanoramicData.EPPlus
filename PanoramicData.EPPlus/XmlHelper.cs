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
 * Author							Change						         Date
 * ******************************************************************************
 * Jan Källman		    Initial Release		         2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 * Eyal Seagull       Add "CreateComplexNode"    2012-04-03
 * Eyal Seagull       Add "DeleteTopNode"        2012-04-13
 *******************************************************************************/

using OfficeOpenXml.Style;
using System.Text;
using System.Xml;
namespace OfficeOpenXml;

/// <summary>
/// Help class containing XML functions. 
/// Can be Inherited 
/// </summary>
public abstract partial class XmlHelper
{
	internal delegate int ChangedEventHandler(StyleBase sender, Style.StyleChangeEventArgs e);

	private protected XmlHelper(XmlNamespaceManager nameSpaceManager)
	{
		TopNode = null;
		NameSpaceManager = nameSpaceManager;
	}

	private protected XmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode? topNode)
	{
		TopNode = topNode;
		NameSpaceManager = nameSpaceManager;
	}

	internal XmlNamespaceManager NameSpaceManager { get; set; }

	internal XmlNode? TopNode { get; set; }

	string[]? _schemaNodeOrder;

	/// <summary>
	/// Schema order list
	/// </summary>
	internal string[]? SchemaNodeOrder
	{
		get
		{
			return _schemaNodeOrder;
		}
		set
		{
			_schemaNodeOrder = value;
		}
	}

	/// <summary>
	/// Options to insert a node in the XmlDocument
	/// </summary>
	internal enum eNodeInsertOrder
	{
		/// <summary>
		/// Insert as first node of "topNode"
		/// </summary>
		First,

		/// <summary>
		/// Insert as the last child of "topNode"
		/// </summary>
		Last,

		/// <summary>
		/// Insert after the "referenceNode"
		/// </summary>
		After,

		/// <summary>
		/// Insert before the "referenceNode"
		/// </summary>
		Before,

		/// <summary>
		/// Use the Schema List to insert in the right order. If the Schema list
		/// is null or empty, consider "Last" as the selected option
		/// </summary>
		SchemaOrder
	}
	internal static Uri GetNewUri(Packaging.ZipPackage package, string sUri)
	{
		var id = 1;
		return GetNewUri(package, sUri, ref id);
	}
	internal static Uri GetNewUri(Packaging.ZipPackage package, string sUri, ref int id)
	{
		Uri uri = new(string.Format(sUri, id), UriKind.Relative);
		while (package.PartExists(uri))
		{
			uri = new Uri(string.Format(sUri, ++id), UriKind.Relative);
		}

		return uri;
	}
	internal static void LoadXmlSafe(XmlDocument xmlDoc, Stream stream)
	{
		XmlReaderSettings settings = new()
		{
			// Disable entity parsing (to avoid XML bombs, External Entity Attacks etc.).
			DtdProcessing = DtdProcessing.Prohibit
		};
		var reader = XmlReader.Create(stream, settings);
		xmlDoc.Load(reader);
	}
	internal static void LoadXmlSafe(XmlDocument xmlDoc, string xml, Encoding encoding)
	{
		var stream = new MemoryStream(encoding.GetBytes(xml));
		LoadXmlSafe(xmlDoc, stream);
	}
}

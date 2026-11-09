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
 *******************************************************************************
 * Jan Källman		Added		25-Oct-2012
 *******************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using ZipCompressionLevel = System.IO.Compression.CompressionLevel;

namespace OfficeOpenXml.Packaging;

/// <summary>
/// Specifies whether the target is inside or outside the System.IO.Packaging.Package.
/// </summary>
public enum TargetMode
{
	/// <summary>
	/// The relationship references a part that is inside the package.
	/// </summary>
	Internal = 0,
	/// <summary>
	/// The relationship references a resource that is external to the package.
	/// </summary>
	External = 1,
}
/// <summary>
/// Represent an OOXML Zip package.
/// </summary>
public class ZipPackage : ZipPackageRelationshipBase
{
	internal class ContentType
	{
		internal string Name;
		internal bool IsExtension;
		internal string Match;
		public ContentType(string name, bool isExtension, string match)
		{
			Name = name;
			IsExtension = isExtension;
			Match = match;
		}
	}
	private const string NotAPackageMessage = "The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor.";

	readonly Dictionary<string, ZipPackagePart> Parts = new(StringComparer.OrdinalIgnoreCase);
	internal Dictionary<string, ContentType> _contentTypes = new(StringComparer.OrdinalIgnoreCase);
	internal char _dirSeparator = '/';
	internal ZipPackage()
	{
		AddNew();
	}

	private void AddNew()
	{
		_contentTypes.Add("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
		_contentTypes.Add("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
	}

	internal ZipPackage(Stream stream)
	{
		if (stream == null || stream.Length == 0)
		{
			AddNew();
		}
		else
		{
			var rels = new Dictionary<string, string>();
			stream.Seek(0, SeekOrigin.Begin);
			using var zip = OpenArchive(stream);
			var hasContentTypeXml = ReadEntries(zip, rels);
			ApplyPartRelationsAndContentTypes(rels);

			if (!hasContentTypeXml)
			{
				throw new InvalidDataException(NotAPackageMessage);
			}
		}
	}

	private static ZipArchive OpenArchive(Stream stream)
	{
		try
		{
			return new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
		}
		catch (InvalidDataException ex)
		{
			throw new InvalidDataException(NotAPackageMessage, ex);
		}
	}

	private bool ReadEntries(ZipArchive zip, Dictionary<string, string> rels)
	{
		if (zip.Entries.Count == 0)
		{
			throw new InvalidDataException(NotAPackageMessage);
		}

		_dirSeparator = zip.Entries[0].FullName.Contains('\\') ? '\\' : '/';

		var hasContentTypeXml = false;
		foreach (var e in zip.Entries)
		{
			if (e.Length > 0)
			{
				var b = new byte[e.Length];
				using (var entryStream = e.Open())
				{
					entryStream.ReadExactly(b);
				}

				hasContentTypeXml |= ReadEntry(e.FullName, b, rels);
			}
		}

		return hasContentTypeXml;
	}

	// Returns true when the entry is the content types part.
	private bool ReadEntry(string fileName, byte[] b, Dictionary<string, string> rels)
	{
		if (fileName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase))
		{
			AddContentTypes(Encoding.UTF8.GetString(b));
			return true;
		}

		if (fileName.Equals($"_rels{_dirSeparator}.rels", StringComparison.OrdinalIgnoreCase))
		{
			ReadRelation(Encoding.UTF8.GetString(b), "");
		}
		else if (fileName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
		{
			rels.Add(GetUriKey(fileName), Encoding.UTF8.GetString(b));
		}
		else
		{
			var part = new ZipPackagePart(this, fileName)
			{
				Stream = new MemoryStream()
			};
			part.Stream.Write(b, 0, b.Length);
			Parts.Add(GetUriKey(fileName), part);
		}

		return false;
	}

	private void ApplyPartRelationsAndContentTypes(Dictionary<string, string> rels)
	{
		foreach (var p in Parts)
		{
			var name = Path.GetFileName(p.Key);
			var extension = Path.GetExtension(p.Key);
			var relFile = string.Format("{0}_rels/{1}.rels", p.Key[..^name.Length], name);
			if (rels.TryGetValue(relFile, out var value))
			{
				p.Value.ReadRelation(value, p.Value.Uri.OriginalString);
			}

			if (_contentTypes.TryGetValue(p.Key, out var value2))
			{
				p.Value.ContentType = value2.Name;
			}
			else if (extension.Length > 1 && _contentTypes.ContainsKey(extension[1..]))
			{
				p.Value.ContentType = _contentTypes[extension[1..]].Name;
			}
		}
	}

	private void AddContentTypes(string xml)
	{
		var doc = new XmlDocument();
		XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

		foreach (XmlElement c in doc.DocumentElement.ChildNodes)
		{
			ContentType ct;
			if (string.IsNullOrEmpty(c.GetAttribute("Extension")))
			{
				ct = new ContentType(c.GetAttribute("ContentType"), false, c.GetAttribute("PartName"));
			}
			else
			{
				ct = new ContentType(c.GetAttribute("ContentType"), true, c.GetAttribute("Extension"));
			}

			_contentTypes.Add(GetUriKey(ct.Match), ct);
		}
	}

	#region Methods
	internal ZipPackagePart CreatePart(Uri partUri, string contentType) => CreatePart(partUri, contentType, CompressionLevel.Default);
	internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel)
	{
		if (PartExists(partUri))
		{
			throw (new InvalidOperationException("Part already exist"));
		}

		var part = new ZipPackagePart(this, partUri, contentType, compressionLevel);
		_contentTypes.Add(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
		Parts.Add(GetUriKey(part.Uri.OriginalString), part);
		return part;
	}
	internal ZipPackagePart GetPart(Uri partUri) => PartExists(partUri)
			? Parts.Single(x => x.Key.Equals(GetUriKey(partUri.OriginalString), StringComparison.OrdinalIgnoreCase)).Value
			: throw (new InvalidOperationException("Part does not exist."));

	internal static string GetUriKey(string uri)
	{
		var ret = uri.Replace('\\', '/');
		if (ret[0] != '/')
		{
			ret = '/' + ret;
		}

		return ret;
	}
	internal bool PartExists(Uri partUri)
	{
		var uriKey = GetUriKey(partUri.OriginalString.ToLowerInvariant());
		return Parts.ContainsKey(uriKey);
	}
	#endregion

	internal void DeletePart(Uri Uri)
	{
		var delList = new List<object[]>();
		foreach (var p in Parts.Values)
		{
			foreach (var r in p.GetRelationships())
			{
				if (UriHelper.ResolvePartUri(p.Uri, r.TargetUri).OriginalString.Equals(Uri.OriginalString, StringComparison.OrdinalIgnoreCase))
				{
					delList.Add([r.Id, p]);
				}
			}
		}

		foreach (var o in delList)
		{
			((ZipPackagePart)o[1]).DeleteRelationship(o[0].ToString());
		}

		var rels = GetPart(Uri).GetRelationships();
		while (rels.Count > 0)
		{
			rels.Remove(rels.First().Id);
		}

		_contentTypes.Remove(GetUriKey(Uri.OriginalString));
		Parts.Remove(GetUriKey(Uri.OriginalString));
	}
	internal void Save(Stream stream)
	{
		// ZipArchive switches to Zip64 only when an archive exceeds the classic limits.
		using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
		var contentTypes = Encoding.UTF8.GetBytes(GetContentTypeXml());
		WriteEntry(archive, "[Content_Types].xml", _compression, entryStream => entryStream.Write(contentTypes, 0, contentTypes.Length));
		_rels.WriteZip(archive, "_rels/.rels", _compression);
		ZipPackagePart ssPart = null;
		foreach (var part in Parts.Values)
		{
			if (part.ContentType != ExcelPackage.contentTypeSharedString)
			{
				part.WriteZip(archive);
			}
			else
			{
				ssPart = part;
			}
		}
		//Shared strings must be saved after all worksheets. The ss dictionary is populated when that workheets are saved (to get the best performance).
		ssPart?.WriteZip(archive);
	}

	internal static void WriteEntry(ZipArchive archive, string name, CompressionLevel compressionLevel, Action<Stream> write)
	{
		var entry = archive.CreateEntry(GetZipEntryName(name), ToZipCompressionLevel(compressionLevel));
		using var entryStream = entry.Open();
		write(entryStream);
	}

	internal static string GetZipEntryName(string uri) => uri.Replace('\\', '/').TrimStart('/');

	// Level0 (None) deflates at the default level, as the DotNetZip writer did.
	internal static ZipCompressionLevel ToZipCompressionLevel(CompressionLevel compressionLevel) => compressionLevel switch
	{
		CompressionLevel.Level1 or CompressionLevel.Level2 or CompressionLevel.Level3 => ZipCompressionLevel.Fastest,
		CompressionLevel.Level9 => ZipCompressionLevel.SmallestSize,
		_ => ZipCompressionLevel.Optimal,
	};

	private string GetContentTypeXml()
	{
		StringBuilder xml = new("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
		foreach (var ct in _contentTypes.Values)
		{
			if (ct.IsExtension)
			{
				xml.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", ct.Name, ct.Match);
			}
			else
			{
				xml.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", ct.Name, GetUriKey(ct.Match));
			}
		}

		xml.Append("</Types>");
		return xml.ToString();
	}
	internal static void Flush()
	{
		// Nothing is buffered: the package is only written by Save
	}
	internal static void Close()
	{
		// Nothing to release: the package holds no open stream between calls
	}
	CompressionLevel _compression = CompressionLevel.Default;
	public CompressionLevel Compression
	{
		get
		{
			return _compression;
		}
		set
		{
			foreach (var part in Parts.Values)
			{
				if (part.CompressionLevel == _compression)
				{
					part.CompressionLevel = value;
				}
			}

			_compression = value;
		}
	}
}

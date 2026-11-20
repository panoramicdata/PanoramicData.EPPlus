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
using System;
using System.IO;
using System.IO.Compression;

namespace OfficeOpenXml.Packaging;

internal class ZipPackagePart : ZipPackageRelationshipBase, IDisposable
{
	internal delegate void SaveHandlerDelegate(Stream stream);

	internal ZipPackagePart(ZipPackage package, string entryName)
	{
		Package = package;
		SaveHandler = null;
		Uri = new Uri(ZipPackage.GetUriKey(entryName), UriKind.Relative);
	}
	internal ZipPackagePart(ZipPackage package, Uri partUri, string contentType, CompressionLevel compressionLevel)
	{
		Package = package;
		Uri = partUri;
		ContentType = contentType;
		CompressionLevel = compressionLevel;
	}
	internal ZipPackage Package { get; set; }
	internal CompressionLevel CompressionLevel;
	MemoryStream _stream;
	internal MemoryStream Stream
	{
		get
		{
			return _stream;
		}
		set
		{
			_stream = value;
		}
	}
	internal override ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
	{

		var rel = base.CreateRelationship(targetUri, targetMode, relationshipType);
		rel.SourceUri = Uri;
		return rel;
	}
	internal MemoryStream GetStream() => GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite);
	internal MemoryStream GetStream(FileMode fileMode) => GetStream(FileMode.Create, FileAccess.ReadWrite);
	internal MemoryStream GetStream(FileMode fileMode, FileAccess fileAccess)
	{
		if (_stream == null || fileMode == FileMode.CreateNew || fileMode == FileMode.Create)
		{
			_stream = new MemoryStream();
		}
		else
		{
			_stream.Seek(0, SeekOrigin.Begin);
		}

		return _stream;
	}

	string _contentType = "";
	public string ContentType
	{
		get
		{
			return _contentType;
		}
		internal set
		{
			if (!string.IsNullOrEmpty(_contentType))
			{
				if (Package._contentTypes.ContainsKey(ZipPackage.GetUriKey(Uri.OriginalString)))
				{
					Package._contentTypes.Remove(ZipPackage.GetUriKey(Uri.OriginalString));
					Package._contentTypes.Add(ZipPackage.GetUriKey(Uri.OriginalString), new ZipPackage.ContentType(value, false, Uri.OriginalString));
				}
			}

			_contentType = value;
		}
	}
	public Uri Uri { get; private set; }
	internal SaveHandlerDelegate SaveHandler
	{
		get;
		set;
	}
	internal void WriteZip(ZipArchive archive)
	{
		if (SaveHandler == null)
		{
			var b = GetStream().ToArray();
			if (b.Length == 0)
			{
				return;
			}

			ZipPackage.WriteEntry(archive, Uri.OriginalString, CompressionLevel, entryStream => entryStream.Write(b, 0, b.Length));
		}
		else
		{
			ZipPackage.WriteEntry(archive, Uri.OriginalString, CompressionLevel, entryStream => SaveHandler(entryStream));
		}

		if (_rels.Count > 0)
		{
			var f = Uri.OriginalString;
			var name = Path.GetFileName(f);
			_rels.WriteZip(archive, string.Format("{0}_rels/{1}.rels", f[..^name.Length], name), CompressionLevel);
		}
	}

	public void Dispose()
	{
		_stream.Close();
		_stream.Dispose();
	}
}

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
 * Jan Källman		Added		10-SEP-2009
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// An image that fills the background of the worksheet.
/// </summary>
public class ExcelBackgroundImage : XmlHelper
{
	readonly ExcelWorksheet _workSheet;
	/// <summary>
	/// 
	/// </summary>
	/// <param name="nsm"></param>
	/// <param name="topNode">The topnode of the worksheet</param>
	/// <param name="workSheet">Worksheet reference</param>
	internal ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet) :
		base(nsm, topNode)
	{
		_workSheet = workSheet;
	}

	const string BACKGROUNDPIC_PATH = "d:picture/@r:id";
	/// <summary>
	/// The background image file's bytes, exactly as stored in the package. Null when there is no background image.
	/// Setting the property stores the bytes unchanged; set it to null to remove the background image.
	/// </summary>
	public byte[] ImageBytes
	{
		get
		{
			var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
			if (!string.IsNullOrEmpty(relID))
			{
				var rel = _workSheet.Part.GetRelationship(relID);
				var imagePart = _workSheet.Part.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
				return imagePart.GetStream().ToArray();
			}

			return null;
		}
		set
		{
			DeletePrevImage();
			if (value == null)
			{
				DeleteAllNode(BACKGROUNDPIC_PATH);
			}
			else
			{
				var contentType = ExcelPicture.GetContentTypeFromBytes(value);
				var uri = GetNewUri(_workSheet._package.Package, "/xl/media/image{0}" + ExcelPicture.GetExtension(contentType));
				var ii = _workSheet.Workbook._package.AddImage(value, uri, contentType);
				var rel = _workSheet.Part.CreateRelationship(ii.Uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
				SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
			}
		}
	}
	/// <summary>
	/// Set the picture from a stream containing an image file. The stream is read to the end.
	/// </summary>
	/// <param name="PictureStream">The image file's stream.</param>
	public void SetFromStream(Stream PictureStream)
	{
		ArgumentNullException.ThrowIfNull(PictureStream);
		using var ms = new MemoryStream();
		PictureStream.CopyTo(ms);
		ImageBytes = ms.ToArray();
	}
	/// <summary>
	/// Set the picture from an image file. 
	/// The image file will be saved as a blob, so make sure Excel supports the image format.
	/// </summary>
	/// <param name="PictureFile">The image file.</param>
	public void SetFromFile(FileInfo PictureFile)
	{
		DeletePrevImage();

		byte[] fileBytes;
		try
		{
			fileBytes = File.ReadAllBytes(PictureFile.FullName);
			ImageHeaderReader.Read(fileBytes);
		}
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException)
		{
			throw (new InvalidDataException("File is not a supported image-file or is corrupt", ex));
		}

		var contentType = ExcelPicture.GetContentType(PictureFile.Extension);
		var imageURI = GetNewUri(_workSheet._package.Package, "/xl/media/" + PictureFile.Name[..^PictureFile.Extension.Length] + "{0}" + PictureFile.Extension);

		var ii = _workSheet.Workbook._package.AddImage(fileBytes, imageURI, contentType);


		if (_workSheet.Part.Package.PartExists(imageURI) && ii.RefCount == 1) //The file exists with another content, overwrite it.
		{
			//Remove the part if it exists
			_workSheet.Part.Package.DeletePart(imageURI);
		}

		var imagePart = _workSheet.Part.Package.CreatePart(imageURI, contentType, CompressionLevel.None);
		//Save the picture to package.

		var strm = imagePart.GetStream(FileMode.Create, FileAccess.Write);
		strm.Write(fileBytes, 0, fileBytes.Length);

		var rel = _workSheet.Part.CreateRelationship(imageURI, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
		SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
	}
	private void DeletePrevImage()
	{
		var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
		if (relID != "")
		{
			var img = ImageBytes;
			var ii = img == null ? null : _workSheet.Workbook._package.GetImageInfo(img);

			//Delete the relation
			_workSheet.Part.DeleteRelationship(relID);

			//Delete the image if there are no other references.
			if (ii != null && ii.RefCount == 1)
			{
				if (_workSheet.Part.Package.PartExists(ii.Uri))
				{
					_workSheet.Part.Package.DeletePart(ii.Uri);
				}
			}

		}
	}
}

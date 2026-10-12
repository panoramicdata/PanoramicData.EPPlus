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
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Utils;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// An image object
/// </summary>
public sealed class ExcelPicture : ExcelDrawing
{
	#region "Constructors"

	internal ExcelPicture(ExcelDrawings drawings, XmlNode node) :
		base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
	{
		var picNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
		if (picNode != null)
		{
			RelPic = drawings.Part.GetRelationship(picNode.Attributes["r:embed"].Value);
			UriPic = UriHelper.ResolvePartUri(drawings.UriDrawing, RelPic.TargetUri);

			Part = drawings.Part.Package.GetPart(UriPic);
			FileInfo f = new(UriPic.OriginalString);
			ContentType = GetContentType(f.Extension);

			// The bytes are hashed as stored. They are not retained: ImageBytes reads them from the package on demand.
			var iby = Part.GetStream().ToArray();
			var ii = _drawings._package.LoadImage(iby, UriPic, Part);
			ImageHash = ii.Hash;

			var relID = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@r:id");
			if (!string.IsNullOrEmpty(relID))
			{
				HypRel = drawings.Part.GetRelationship(relID);
				if (HypRel.TargetUri.IsAbsoluteUri)
				{
					Hyperlink = new ExcelHyperLink(HypRel.TargetUri.AbsoluteUri);
				}
				else
				{
					Hyperlink = new ExcelHyperLink(HypRel.TargetUri.OriginalString, UriKind.Relative);
				}

				((ExcelHyperLink)Hyperlink).ToolTip = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@tooltip");
			}
		}
	}

	/// <summary>
	/// Creates a picture from an image file's bytes. The size and resolution are read from the image header.
	/// The bytes are written into the package verbatim.
	/// </summary>
	internal ExcelPicture(ExcelDrawings drawings, XmlNode node, byte[] imageBytes, string uriTemplate, Uri hyperlink) :
		base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
	{
		var info = ImageHeaderReader.Read(imageBytes);
		Create(node, imageBytes, info.ContentType, uriTemplate, info.Width, info.Height, info.HorizontalResolution, info.VerticalResolution, hyperlink);
	}

	/// <summary>
	/// Creates a picture from an image file's bytes, with the display size supplied by the caller.
	/// No header is read. The bytes are written into the package verbatim, with the content type
	/// sniffed from the byte signature.
	/// </summary>
	internal ExcelPicture(ExcelDrawings drawings, XmlNode node, byte[] imageBytes, int widthPixels, int heightPixels, Uri hyperlink) :
		base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
	{
		var contentType = GetContentTypeFromBytes(imageBytes);
		Create(node, imageBytes, contentType, "/xl/media/image{0}" + GetExtension(contentType), widthPixels, heightPixels, ImageHeaderInfo.DefaultDpi, ImageHeaderInfo.DefaultDpi, hyperlink);
	}

	private void Create(XmlNode node, byte[] imageBytes, string contentType, string uriTemplate, int widthPixels, int heightPixels, float horizontalDpi, float verticalDpi, Uri hyperlink)
	{
		var picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
		node.InsertAfter(picNode, node.SelectSingleNode("xdr:to", NameSpaceManager));
		Hyperlink = hyperlink;
		picNode.InnerXml = PicStartXml();

		node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);

		var drawings = _drawings;
		var package = drawings.Worksheet._package.Package;
		_imageBytes = imageBytes;
		ContentType = contentType;
		UriPic = GetNewUri(package, uriTemplate);
		var ii = drawings._package.AddImage(imageBytes, UriPic, ContentType);
		string relID;
		if (!drawings._hashes.TryGetValue(ii.Hash, out var value))
		{
			Part = ii.Part;
			RelPic = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, ii.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
			relID = RelPic.Id;
			drawings._hashes.Add(ii.Hash, relID);
		}
		else
		{
			relID = value;
			var rel = drawings.Part.GetRelationship(relID);
			UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
		}

		ImageHash = ii.Hash;
		_height = heightPixels;
		_width = widthPixels;
		EditAs = eEditAs.OneCell;
		SetPixelWidth(widthPixels, horizontalDpi);
		SetPixelHeight(heightPixels, verticalDpi);
		//Create relationship
		node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
		Packaging.ZipPackage.Flush();
	}

	/// <summary>
	/// Content type from the byte signature, so no decode and no file extension is needed.
	/// </summary>
	internal static string GetContentTypeFromBytes(byte[] imageBytes) => ImageHeaderReader.GetContentType(imageBytes);

	internal static string GetExtension(string contentType) => contentType switch
	{
		"image/png" => ".png",
		"image/gif" => ".gif",
		"image/bmp" => ".bmp",
		"image/x-tiff" => ".tif",
		"image/x-emf" => ".emf",
		"image/x-wmf" => ".wmf",
		_ => ".jpg",
	};

	internal static string GetContentType(string extension) => extension.ToLower(CultureInfo.InvariantCulture) switch
	{
		".bmp" => "image/bmp",
		".jpg" or ".jpeg" => "image/jpeg",
		".gif" => "image/gif",
		".png" => "image/png",
		".cgm" => "image/cgm",
		".emf" => "image/x-emf",
		".eps" => "image/x-eps",
		".pcx" => "image/x-pcx",
		".tga" => "image/x-tga",
		".tif" or ".tiff" => "image/x-tiff",
		".wmf" => "image/x-wmf",
		_ => "image/jpeg",
	};
	#endregion
	private string SavePicture(byte[] imageBytes)
	{
		var contentType = GetContentTypeFromBytes(imageBytes);
		var ii = _drawings._package.AddImage(imageBytes, GetNewUri(_drawings.Worksheet._package.Package, "/xl/media/image{0}" + GetExtension(contentType)), contentType);

		ImageHash = ii.Hash;
		if (_drawings._hashes.TryGetValue(ii.Hash, out var relID))
		{
			var rel = _drawings.Part.GetRelationship(relID);
			UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
			return relID;
		}
		else
		{
			UriPic = ii.Uri;
			ImageHash = ii.Hash;
		}

		ContentType = contentType;
		RelPic = _drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(_drawings.UriDrawing, UriPic), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

		_drawings._hashes.Add(ii.Hash, RelPic.Id);

		return RelPic.Id;
	}

	private string PicStartXml()
	{
		StringBuilder xml = new();

		xml.Append("<xdr:nvPicPr>");

		if (Hyperlink == null)
		{
			xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
		}
		else
		{
			HypRel = _drawings.Part.CreateRelationship(Hyperlink, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
			xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\">", _id);
			if (HypRel != null)
			{
				if (Hyperlink is ExcelHyperLink link)
				{
					xml.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" tooltip=\"{1}\"/>",
					  HypRel.Id, link.ToolTip);
				}
				else
				{
					xml.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" />",
					  HypRel.Id);
				}
			}

			xml.Append("</xdr:cNvPr>");
		}

		xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

		return xml.ToString();
	}

	internal string ImageHash { get; set; }
	byte[] _imageBytes = null;
	/// <summary>
	/// The image file's bytes, exactly as stored in the package (PNG, JPEG, GIF, BMP, TIFF, EMF or WMF).
	/// Setting it replaces the picture; the bytes are written to the package verbatim.
	/// </summary>
	public byte[] ImageBytes
	{
		get
		{
			if (_imageBytes != null)
			{
				return _imageBytes;
			}

			return Part?.GetStream().ToArray();
		}
		set
		{
			if (value != null)
			{
				_imageBytes = value;
				try
				{
					var relID = SavePicture(value);

					//Create relationship
					TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
				}
				catch (Exception ex)
				{
					throw new Exception("Can't save image - " + ex.Message, ex);
				}
			}
		}
	}
	/// <summary>
	/// The image's MIME content type, for example image/png or image/jpeg.
	/// </summary>
	public string ContentType
	{
		get;
		internal set;
	}
	/// <summary>
	/// Set the size of the image in percent from the orginal size
	/// Note that resizing columns / rows after using this function will effect the size of the picture
	/// </summary>
	/// <param name="Percent">Percent</param>
	public override void SetSize(int Percent)
	{
		var bytes = ImageBytes;
		if (bytes == null)
		{
			base.SetSize(Percent);
		}
		else
		{
			var info = ImageHeaderReader.Read(bytes);
			_width = info.Width;
			_height = info.Height;

			_width = (int)(_width * ((decimal)Percent / 100));
			_height = (int)(_height * ((decimal)Percent / 100));

			SetPixelWidth(_width, info.HorizontalResolution);
			SetPixelHeight(_height, info.VerticalResolution);
		}
	}
	internal Uri UriPic { get; set; }
	internal Packaging.ZipPackageRelationship RelPic { get; set; }
	internal Packaging.ZipPackageRelationship HypRel { get; set; }
	internal Packaging.ZipPackagePart Part;

	internal new string Id => Name;

	ExcelDrawingFill? _fill = null;
	/// <summary>
	/// Fill
	/// </summary>
	public ExcelDrawingFill Fill
	{
		get
		{
			_fill ??= new ExcelDrawingFill(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr");
			return _fill;
		}
	}
	ExcelDrawingBorder? _border = null;

	/// <summary>
	/// Border
	/// </summary>
	public ExcelDrawingBorder Border
	{
		get
		{
			_border ??= new ExcelDrawingBorder(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr/a:ln");
			return _border;
		}
	}

	/// <summary>
	/// Hyperlink
	/// </summary>
	public Uri? Hyperlink { get; private set; }

	internal override void DeleteMe()
	{
		_drawings._package.RemoveImage(ImageHash);
		base.DeleteMe();
	}
	public override void Dispose()
	{
		base.Dispose();
		Hyperlink = null;
		_imageBytes = null;
	}
}
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// How a picture will be aligned in the header/footer
/// </summary>
public enum PictureAlignment
{
	/// <summary>
	/// The picture will be added to the left aligned text
	/// </summary>
	Left,
	/// <summary>
	/// The picture will be added to the centered text
	/// </summary>
	Centered,
	/// <summary>
	/// The picture will be added to the right aligned text
	/// </summary>
	Right
}
/// <summary>
/// Print header and footer 
/// </summary>
[SuppressMessage("Minor Code Smell", "S1104:Fields should not have public accessibility", Justification = "Public API: the text fields are published; converting them to properties would break binary compatibility.")]
[SuppressMessage("Minor Code Smell", "S2357:Fields should be private", Justification = "Public API: the text fields are published; converting them to properties would break binary compatibility.")]
public class ExcelHeaderFooterText
{
	readonly ExcelWorksheet _ws;
	readonly string _hf;
	internal ExcelHeaderFooterText(XmlNode TextNode, ExcelWorksheet ws, string hf)
	{
		_ws = ws;
		_hf = hf;
		if (TextNode == null || string.IsNullOrEmpty(TextNode.InnerText))
		{
			return;
		}

		var text = TextNode.InnerText;
		var code = text[..2];
		var startPos = 2;
		var pos = startPos;
		while (pos < text.Length - 2)
		{
			var newCode = text.Substring(pos, 2);
			if (newCode is "&C" or "&R")
			{
				SetText(code, text[startPos..pos]);
				startPos = pos + 2;
				pos = startPos;
				code = newCode;
			}

			pos++;
		}

		SetText(code, text[startPos..]);
	}
	private void SetText(string code, string text)
	{
		switch (code)
		{
			case "&L":
				LeftAlignedText = text;
				break;
			case "&C":
				CenteredText = text;
				break;
			default:
				RightAlignedText = text;
				break;
		}
	}
	/// <summary>
	/// Get/set the text to appear on the left hand side of the header (or footer) on the worksheet.
	/// </summary>
	public string LeftAlignedText;
	/// <summary>
	/// Get/set the text to appear in the center of the header (or footer) on the worksheet.
	/// </summary>
	public string CenteredText;
	/// <summary>
	/// Get/set the text to appear on the right hand side of the header (or footer) on the worksheet.
	/// </summary>
	public string RightAlignedText;
	/// <summary>
	/// Inserts a picture at the end of the text in the header or footer
	/// </summary>
	/// <param name="Picture">The image file's bytes (PNG, JPEG, GIF, BMP, TIFF, EMF or WMF)</param>
	/// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
	public ExcelVmlDrawingPicture InsertPicture(byte[] Picture, PictureAlignment Alignment)
	{
		ArgumentNullException.ThrowIfNull(Picture);
		// Read the header first: ValidateImage appends the picture placeholder to the section text,
		// which must not happen for data that is rejected.
		var header = ReadHeader(Picture);
		var id = ValidateImage(Alignment);
		var contentType = header.ContentType;
		var uriPic = XmlHelper.GetNewUri(_ws._package.Package, "/xl/media/image{0}" + ExcelPicture.GetExtension(contentType));
		var ii = _ws.Workbook._package.AddImage(Picture, uriPic, contentType);

		return AddImage(header, id, ii);
	}
	/// <summary>
	/// Inserts a picture at the end of the text in the header or footer
	/// </summary>
	/// <param name="Picture">A stream containing the image file. It is read to the end.</param>
	/// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
	public ExcelVmlDrawingPicture InsertPicture(Stream Picture, PictureAlignment Alignment)
	{
		ArgumentNullException.ThrowIfNull(Picture);
		using var ms = new MemoryStream();
		Picture.CopyTo(ms);
		return InsertPicture(ms.ToArray(), Alignment);
	}
	/// <summary>
	/// Inserts a picture at the end of the text in the header or footer
	/// </summary>
	/// <param name="PictureFile">The image file</param>
	/// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
	public ExcelVmlDrawingPicture InsertPicture(FileInfo PictureFile, PictureAlignment Alignment)
	{
		ArgumentNullException.ThrowIfNull(PictureFile);
		byte[] imgBytes;
		ImageHeaderInfo header;
		try
		{
			if (!PictureFile.Exists)
			{
				throw new FileNotFoundException(string.Format("{0} is missing", PictureFile.FullName));
			}

			imgBytes = File.ReadAllBytes(PictureFile.FullName);
			header = ReadHeader(imgBytes);
		}
		catch (Exception ex)
		{
			throw new InvalidDataException("File is not a supported image-file or is corrupt", ex);
		}

		var id = ValidateImage(Alignment);
		var contentType = ExcelPicture.GetContentType(PictureFile.Extension);
		var uriPic = XmlHelper.GetNewUri(_ws._package.Package, "/xl/media/" + PictureFile.Name[..^PictureFile.Extension.Length] + "{0}" + PictureFile.Extension);

		var ii = _ws.Workbook._package.AddImage(imgBytes, uriPic, contentType);

		return AddImage(header, id, ii);
	}

	private static ImageHeaderInfo ReadHeader(byte[] bytes) => ImageHeaderReader.Read(bytes);

	private ExcelVmlDrawingPicture AddImage(ImageHeaderInfo Picture, string id, ExcelPackage.ImageInfo ii)
	{
		double width = Picture.Width * 72 / Picture.HorizontalResolution,      //Pixel --> Points
			   height = Picture.Height * 72 / Picture.VerticalResolution;      //Pixel --> Points
																			   //Add VML-drawing            
		return _ws.HeaderFooter.Pictures.Add(id, ii.Uri, "", width, height);
	}
	private string ValidateImage(PictureAlignment Alignment)
	{
		var id = string.Concat(Alignment.ToString()[0], _hf);
		foreach (ExcelVmlDrawingPicture image in _ws.HeaderFooter.Pictures)
		{
			if (image.Id == id)
			{
				throw new InvalidOperationException("A picture already exists in this section");
			}
		}
		//Add the image placeholder to the end of the text
		switch (Alignment)
		{
			case PictureAlignment.Left:
				LeftAlignedText += ExcelHeaderFooter.Image;
				break;
			case PictureAlignment.Centered:
				CenteredText += ExcelHeaderFooter.Image;
				break;
			default:
				RightAlignedText += ExcelHeaderFooter.Image;
				break;
		}

		return id;
	}
}

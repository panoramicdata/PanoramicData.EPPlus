using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.IO;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// MagicSuite MS-25871: the byte-based AddPicture overloads, which work without any graphics library
/// on every platform. Nothing in this class may reference System.Drawing image types.
/// </summary>
[TestClass]
public class AddPictureFromBytesTests
{
	// A 1x1 PNG - real image bytes with a real PNG signature, produced without GDI+.
	private static readonly byte[] OnePixelPng = Convert.FromBase64String(
		"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

	[TestMethod]
	public void AddPictureFromBytes_RoundTripsThroughASavedPackage()
	{
		byte[] saved;
		using (var package = new ExcelPackage())
		{
			var worksheet = package.Workbook.Worksheets.Add("Pictures");
			var picture = worksheet.Drawings.AddPicture("CatBytes", OnePixelPng, 96, 48);

			Assert.AreEqual("image/png", picture.ContentType);
			Assert.IsTrue(picture.UriPic.OriginalString.EndsWith(".png"), "the part name carries the sniffed extension");

			saved = package.GetAsByteArray();
		}


		using var reopened = new ExcelPackage(new MemoryStream(saved));
		var drawing = reopened.Workbook.Worksheets[0].Drawings["CatBytes"];
		Assert.IsInstanceOfType(drawing, typeof(ExcelPicture));

		var reopenedPicture = (ExcelPicture)drawing;
		CollectionAssert.AreEqual(OnePixelPng, reopenedPicture.ImageBytes, "the picture reads back through the public API");
		using var partStream = reopenedPicture.Part.GetStream();
		using var ms = new MemoryStream();
		partStream.CopyTo(ms);
		CollectionAssert.AreEqual(OnePixelPng, ms.ToArray(), "the bytes must land in the package verbatim");
	}

	[TestMethod]
	public void AddPictureFromBytes_SniffsTheContentTypeFromTheSignature()
	{
		Assert.AreEqual("image/png", ExcelPicture.GetContentTypeFromBytes(OnePixelPng));
		Assert.AreEqual("image/jpeg", ExcelPicture.GetContentTypeFromBytes([0xFF, 0xD8, 0xFF, 0xE0, 1, 2, 3, 4]));
		Assert.AreEqual("image/gif", ExcelPicture.GetContentTypeFromBytes([0x47, 0x49, 0x46, 0x38]));
		Assert.AreEqual("image/bmp", ExcelPicture.GetContentTypeFromBytes([0x42, 0x4D, 0, 0]));
		Assert.AreEqual("image/jpeg", ExcelPicture.GetContentTypeFromBytes([1, 2, 3, 4]), "unknown signatures fall back to jpeg");
	}

	[TestMethod]
	public void AddPictureFromBytes_RejectsBadArguments()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("Pictures");

		Assert.ThrowsException<ArgumentException>(() => worksheet.Drawings.AddPicture("P", null, 10, 10));
		Assert.ThrowsException<ArgumentException>(() => worksheet.Drawings.AddPicture("P", Array.Empty<byte>(), 10, 10));
		Assert.ThrowsException<ArgumentOutOfRangeException>(() => worksheet.Drawings.AddPicture("P", OnePixelPng, 0, 10));
		Assert.ThrowsException<ArgumentOutOfRangeException>(() => worksheet.Drawings.AddPicture("P", OnePixelPng, 10, -1));
	}

	[TestMethod]
	public void AddPictureFromBytes_DuplicateBytes_ShareOnePackagePart()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("Pictures");

		var first = worksheet.Drawings.AddPicture("First", OnePixelPng, 10, 10);
		var second = worksheet.Drawings.AddPicture("Second", OnePixelPng, 20, 20);

		Assert.AreEqual(first.UriPic.OriginalString, second.UriPic.OriginalString, "identical bytes deduplicate to one image part");
	}

	[TestMethod]
	public void AddPictureFromBytes_ReadsSizeFromTheHeader()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("Pictures");

		var picture = worksheet.Drawings.AddPicture("Auto", OnePixelPng);

		Assert.AreEqual("image/png", picture.ContentType);
		Assert.AreEqual(1, picture.GetPixelWidth());
		Assert.AreEqual(1, picture.GetPixelHeight());
	}

	[TestMethod]
	public void AddPictureFromStream_StoresTheBytesVerbatim()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("Pictures");

		var picture = worksheet.Drawings.AddPicture("Streamed", new MemoryStream(OnePixelPng));

		CollectionAssert.AreEqual(OnePixelPng, picture.ImageBytes);
	}

	[TestMethod]
	public void AddPictureFromFile_ReadsBytesAndSize()
	{
		var file = new FileInfo(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png"));
		File.WriteAllBytes(file.FullName, OnePixelPng);
		try
		{
			using var package = new ExcelPackage();
			var worksheet = package.Workbook.Worksheets.Add("Pictures");

			var picture = worksheet.Drawings.AddPicture("File", file);

			CollectionAssert.AreEqual(OnePixelPng, picture.ImageBytes);
			Assert.AreEqual(1, picture.GetPixelWidth());
		}
		finally
		{
			file.Delete();
		}
	}

	[TestMethod]
	public void ImageHeaderReader_ReadsPngJpegGifAndBmp()
	{
		var png = OfficeOpenXml.Compatibility.ImageHeaderReader.Read(OnePixelPng);
		Assert.AreEqual((1, 1), (png.Width, png.Height));

		var jpeg = OfficeOpenXml.Compatibility.ImageHeaderReader.Read(EPPlusTest.Properties.Resources.Test1);
		Assert.AreEqual("image/jpeg", jpeg.ContentType);
		Assert.IsTrue(jpeg.Width > 1 && jpeg.Height > 1);

		// GIF89a, 3x2
		byte[] gif = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 3, 0, 2, 0, 0, 0, 0];
		var g = OfficeOpenXml.Compatibility.ImageHeaderReader.Read(gif);
		Assert.AreEqual((3, 2), (g.Width, g.Height));

		// BITMAPINFOHEADER, 5x4, 2835 pixels per metre (72 DPI)
		var bmp = new byte[54];
		bmp[0] = 0x42;
		bmp[1] = 0x4D;
		bmp[14] = 40;
		bmp[18] = 5;
		bmp[22] = 4;
		bmp[38] = 0x13;
		bmp[39] = 0x0B;
		var b = OfficeOpenXml.Compatibility.ImageHeaderReader.Read(bmp);
		Assert.AreEqual((5, 4), (b.Width, b.Height));
		Assert.AreEqual(72F, b.HorizontalResolution, 0.1F);
	}

	[TestMethod]
	public void ImageHeaderReader_RejectsTruncatedData()
	{
		Assert.ThrowsException<InvalidDataException>(() => OfficeOpenXml.Compatibility.ImageHeaderReader.Read([0x89, 0x50, 0x4E, 0x47]));
	}
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.IO;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// MagicSuite MS-25871: the byte-based AddPicture overload, which must work without System.Drawing
/// so pictures can be added on platforms where GDI+ is unavailable (System.Drawing.Common 7 and
/// later is Windows-only). Nothing in this class may reference System.Drawing.
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
}

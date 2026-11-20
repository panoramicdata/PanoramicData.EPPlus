using OfficeOpenXml;
using OfficeOpenXml.Compatibility;

namespace PanoramicData.EPPlus.Test;

public partial class ImageHeaderReaderTests
{
	[TestMethod]
	public void Emf_SizeIsTheFrameConvertedToPixelsAt96Dpi()
	{
		// 2540 x 1270 hundredths of a millimetre is 1 x 0.5 inch.
		var info = ImageHeaderReader.Read(BuildEmf(2540, 1270));

		Assert.AreEqual("image/x-emf", info.ContentType);
		Assert.AreEqual(96, info.Width);
		Assert.AreEqual(48, info.Height);
		Assert.AreEqual(96F, info.HorizontalResolution);
	}

	[TestMethod]
	public void Emf_Truncated_IsRejectedOnceTheSignatureIsLost()
	{
		var emf = BuildEmf(2540, 1270);

		// 44 bytes hold the frame and the signature, so that is the shortest readable EMF.
		Assert.AreEqual(96, ImageHeaderReader.Read(emf[..44]).Width);

		// Any shorter is no longer recognisable as an EMF and is not a valid image of any kind.
		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(emf[..40]));
	}

	[TestMethod]
	[DataRow((ushort)1440)]
	[DataRow((ushort)0)] // 0 means the default of 1440 units per inch
	public void Wmf_SizeIsTheBoundingBoxConvertedToPixelsAt96Dpi(ushort unitsPerInch)
	{
		var info = ImageHeaderReader.Read(BuildWmf(1440, 720, unitsPerInch));

		Assert.AreEqual("image/x-wmf", info.ContentType);
		Assert.AreEqual(96, info.Width);
		Assert.AreEqual(48, info.Height);
	}

	[TestMethod]
	public void Wmf_CustomUnitsPerInch_ScalesTheSize()
	{
		var info = ImageHeaderReader.Read(BuildWmf(200, 100, 100));

		Assert.AreEqual(192, info.Width);
		Assert.AreEqual(96, info.Height);
	}

	[TestMethod]
	public void Wmf_Truncated_IsRejected()
	{
		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(BuildWmf(1440, 720, 1440)[..10]));
	}

	[TestMethod]
	public void NullBytes_AreRejected()
	{
		Assert.ThrowsExactly<ArgumentNullException>(() => ImageHeaderReader.Read(null));
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_FromBytes_StoresTheBytesVerbatimAndSizesInPoints()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("HF");

		var picture = ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng, PictureAlignment.Centered);

		// 1 pixel at 96 dpi is 0.75 points.
		Assert.AreEqual(0.75, picture.Width, 0.0001);
		Assert.AreEqual(0.75, picture.Height, 0.0001);
		CollectionAssert.AreEqual(OnePixelPng, picture.ImageBytes, "the bytes must be stored unchanged");
		Assert.IsTrue(ws.HeaderFooter.OddHeader.CenteredText.EndsWith(ExcelHeaderFooter.Image));
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_FromStream_MatchesTheByteOverload()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("HF");

		using var stream = new MemoryStream(OnePixelPng);
		var picture = ws.HeaderFooter.OddFooter.InsertPicture(stream, PictureAlignment.Right);

		Assert.AreEqual(0.75, picture.Width, 0.0001);
		CollectionAssert.AreEqual(OnePixelPng, picture.ImageBytes);
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_FromFile_StoresTheFileBytesUnchanged()
	{
		var path = Path.Combine(Path.GetTempPath(), $"hf-{Guid.NewGuid():N}.png");
		File.WriteAllBytes(path, OnePixelPng);
		try
		{
			using var package = new ExcelPackage();
			var ws = package.Workbook.Worksheets.Add("HF");

			var picture = ws.HeaderFooter.FirstHeader.InsertPicture(new FileInfo(path), PictureAlignment.Left);

			Assert.AreEqual(0.75, picture.Width, 0.0001);
			CollectionAssert.AreEqual(OnePixelPng, picture.ImageBytes);
		}
		finally
		{
			File.Delete(path);
		}
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_SecondPictureInTheSameSection_IsRejected()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("HF");
		ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng, PictureAlignment.Left);

		Assert.ThrowsExactly<InvalidOperationException>(() => ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng, PictureAlignment.Left));

		// A different alignment in the same header is a different section.
		var other = ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng, PictureAlignment.Right);
		Assert.IsNotNull(other);
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_InvalidOrTruncatedData_IsRejectedWithoutChangingTheText()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("HF");
		ws.HeaderFooter.OddHeader.LeftAlignedText = "before";

		Assert.ThrowsExactly<InvalidDataException>(() => ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng[..12], PictureAlignment.Left));
		Assert.ThrowsExactly<InvalidDataException>(() => ws.HeaderFooter.OddHeader.InsertPicture(new FileInfo(Path.Combine(Path.GetTempPath(), "does-not-exist.png")), PictureAlignment.Left));

		Assert.AreEqual("before", ws.HeaderFooter.OddHeader.LeftAlignedText, "rejected data must not leave a picture placeholder behind");
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_NullArguments_AreRejected()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("HF");

		Assert.ThrowsExactly<ArgumentNullException>(() => ws.HeaderFooter.OddHeader.InsertPicture((byte[])null, PictureAlignment.Left));
		Assert.ThrowsExactly<ArgumentNullException>(() => ws.HeaderFooter.OddHeader.InsertPicture((Stream)null, PictureAlignment.Left));
		Assert.ThrowsExactly<ArgumentNullException>(() => ws.HeaderFooter.OddHeader.InsertPicture((FileInfo)null, PictureAlignment.Left));
	}

	[TestMethod]
	public void AddPicture_NullFileInfo_ThrowsArgumentNullException()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("Pictures");

		Assert.ThrowsExactly<ArgumentNullException>(() => ws.Drawings.AddPicture("P", (FileInfo)null, null));
	}
}

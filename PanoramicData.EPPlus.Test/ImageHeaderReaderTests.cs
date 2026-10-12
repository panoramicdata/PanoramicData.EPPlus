using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Compatibility;
using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// The managed image header reader, which replaces GDI+ and must behave identically on every platform:
/// TIFF, EMF and WMF dimensions and resolution, malformed and truncated data, and the header/footer
/// InsertPicture overloads that sit on top of it.
/// </summary>
[TestClass]
public class ImageHeaderReaderTests
{
	// A 1x1 PNG with no pHYs chunk, so it reads as 96 dpi.
	private static readonly byte[] OnePixelPng = Convert.FromBase64String(
		"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

	private const ushort Short = 3;
	private const ushort Long = 4;
	private const ushort Rational = 5;
	private const ushort Double = 12;

	private readonly record struct TiffEntry(ushort Tag, ushort Type, uint Value, (uint Num, uint Den)? Fraction = null);

	/// <summary>Builds a single-IFD TIFF. A rational entry's data is placed after the IFD.</summary>
	private static byte[] BuildTiff(bool little, params TiffEntry[] entries)
	{
		void W16(Span<byte> s, ushort v)
		{
			if (little)
			{
				BinaryPrimitives.WriteUInt16LittleEndian(s, v);
			}
			else
			{
				BinaryPrimitives.WriteUInt16BigEndian(s, v);
			}
		}

		void W32(Span<byte> s, uint v)
		{
			if (little)
			{
				BinaryPrimitives.WriteUInt32LittleEndian(s, v);
			}
			else
			{
				BinaryPrimitives.WriteUInt32BigEndian(s, v);
			}
		}

		var ifdLength = 2 + (entries.Length * 12) + 4;
		var buffer = new List<byte>();
		var data = new byte[8 + ifdLength];
		data[0] = data[1] = (byte)(little ? 0x49 : 0x4D);
		W16(data.AsSpan(2), 42);
		W32(data.AsSpan(4), 8);
		W16(data.AsSpan(8), (ushort)entries.Length);

		var rationalAt = data.Length;
		for (var i = 0; i < entries.Length; i++)
		{
			var e = data.AsSpan(10 + (i * 12));
			W16(e, entries[i].Tag);
			W16(e[2..], entries[i].Type);
			W32(e[4..], 1);
			if (entries[i].Fraction is { } f)
			{
				W16(e[2..], Rational);
				W32(e[8..], (uint)(rationalAt + buffer.Count));
				var pair = new byte[8];
				W32(pair, f.Num);
				W32(pair.AsSpan(4), f.Den);
				buffer.AddRange(pair);
			}
			else if (entries[i].Type == Short)
			{
				W16(e[8..], (ushort)entries[i].Value);
			}
			else
			{
				W32(e[8..], entries[i].Value);
			}
		}

		return [.. data, .. buffer];
	}

	private static byte[] BuildEmf(int frameWidth, int frameHeight)
	{
		var b = new byte[88];
		BinaryPrimitives.WriteInt32LittleEndian(b, 1);                  // EMR_HEADER
		BinaryPrimitives.WriteInt32LittleEndian(b.AsSpan(4), 88);
		BinaryPrimitives.WriteInt32LittleEndian(b.AsSpan(32), frameWidth);   // rclFrame.right (0.01 mm)
		BinaryPrimitives.WriteInt32LittleEndian(b.AsSpan(36), frameHeight);  // rclFrame.bottom
		BinaryPrimitives.WriteInt32LittleEndian(b.AsSpan(40), 0x464D4520);   // " EMF"
		return b;
	}

	private static byte[] BuildWmf(short right, short bottom, ushort unitsPerInch)
	{
		var b = new byte[22];
		b[0] = 0xD7;
		b[1] = 0xCD;
		b[2] = 0xC6;
		b[3] = 0x9A;
		BinaryPrimitives.WriteInt16LittleEndian(b.AsSpan(10), right);
		BinaryPrimitives.WriteInt16LittleEndian(b.AsSpan(12), bottom);
		BinaryPrimitives.WriteUInt16LittleEndian(b.AsSpan(14), unitsPerInch);
		return b;
	}

	[DataTestMethod]
	[DataRow(true)]
	[DataRow(false)]
	public void Tiff_ReadsSizeAndResolutionInBothByteOrders(bool little)
	{
		var tiff = BuildTiff(
			little,
			new TiffEntry(256, Short, 200),
			new TiffEntry(257, Long, 100),
			new TiffEntry(282, Rational, 0, (300, 1)),
			new TiffEntry(283, Rational, 0, (150, 1)),
			new TiffEntry(296, Short, 2));

		var info = ImageHeaderReader.Read(tiff);

		Assert.AreEqual("image/x-tiff", info.ContentType);
		Assert.AreEqual(200, info.Width);
		Assert.AreEqual(100, info.Height);
		Assert.AreEqual(300F, info.HorizontalResolution);
		Assert.AreEqual(150F, info.VerticalResolution);
	}

	[TestMethod]
	public void Tiff_CentimetreResolutionUnitIsConvertedToDotsPerInch()
	{
		var tiff = BuildTiff(
			true,
			new TiffEntry(256, Short, 10),
			new TiffEntry(257, Short, 10),
			new TiffEntry(282, Rational, 0, (100, 1)),
			new TiffEntry(283, Rational, 0, (100, 1)),
			new TiffEntry(296, Short, 3));

		var info = ImageHeaderReader.Read(tiff);

		Assert.AreEqual(254F, info.HorizontalResolution, 0.01F);
		Assert.AreEqual(254F, info.VerticalResolution, 0.01F);
	}

	[TestMethod]
	public void Tiff_WithoutResolution_FallsBackToTheDefaultDpi()
	{
		var tiff = BuildTiff(true, new TiffEntry(256, Short, 8), new TiffEntry(257, Short, 4));

		var info = ImageHeaderReader.Read(tiff);

		Assert.AreEqual(96F, info.HorizontalResolution);
		Assert.AreEqual(96F, info.VerticalResolution);
	}

	[TestMethod]
	public void Tiff_UnsupportedTypeOnAnOptionalField_IsIgnored()
	{
		var tiff = BuildTiff(
			true,
			new TiffEntry(256, Short, 8),
			new TiffEntry(257, Short, 4),
			new TiffEntry(282, Double, 0));

		var info = ImageHeaderReader.Read(tiff);

		Assert.AreEqual(8, info.Width);
		Assert.AreEqual(96F, info.HorizontalResolution, "an unreadable resolution falls back to the default");
	}

	[TestMethod]
	public void Tiff_UnsupportedTypeOnARequiredField_IsRejected()
	{
		var tiff = BuildTiff(true, new TiffEntry(256, Double, 8), new TiffEntry(257, Short, 4));

		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(tiff));
	}

	[TestMethod]
	public void Tiff_WithoutWidthAndHeight_IsRejected()
	{
		var tiff = BuildTiff(true, new TiffEntry(282, Rational, 0, (72, 1)));

		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(tiff));
	}

	[TestMethod]
	public void Tiff_IgnoresTagsItDoesNotUse()
	{
		var tiff = BuildTiff(
			false,
			new TiffEntry(254, Long, 0),
			new TiffEntry(256, Short, 6),
			new TiffEntry(257, Short, 3),
			new TiffEntry(262, Short, 2));

		var info = ImageHeaderReader.Read(tiff);

		Assert.AreEqual(6, info.Width);
		Assert.AreEqual(3, info.Height);
	}

	[TestMethod]
	public void Tiff_TruncatedOrWithWildOffsets_IsRejectedNotCrashed()
	{
		var valid = BuildTiff(true, new TiffEntry(256, Short, 8), new TiffEntry(257, Short, 4), new TiffEntry(282, Rational, 0, (72, 1)));

		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(valid[..10]), "cut inside the IFD");
		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(valid[..^4]), "cut inside the rational data");

		var wildIfd = (byte[])valid.Clone();
		BinaryPrimitives.WriteUInt32LittleEndian(wildIfd.AsSpan(4), 0xFFFFFFF0);
		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(wildIfd), "IFD offset beyond the data");

		var wildRational = (byte[])valid.Clone();
		BinaryPrimitives.WriteUInt32LittleEndian(wildRational.AsSpan(10 + (2 * 12) + 8), 0x7FFFFFF0);
		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(wildRational), "rational offset beyond the data");
	}

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
		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(emf[..40]));
	}

	[DataTestMethod]
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
		Assert.ThrowsException<InvalidDataException>(() => ImageHeaderReader.Read(BuildWmf(1440, 720, 1440)[..10]));
	}

	[TestMethod]
	public void NullBytes_AreRejected()
	{
		Assert.ThrowsException<ArgumentNullException>(() => ImageHeaderReader.Read(null));
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

		Assert.ThrowsException<InvalidOperationException>(() => ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng, PictureAlignment.Left));

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

		Assert.ThrowsException<InvalidDataException>(() => ws.HeaderFooter.OddHeader.InsertPicture(OnePixelPng[..12], PictureAlignment.Left));
		Assert.ThrowsException<InvalidDataException>(() => ws.HeaderFooter.OddHeader.InsertPicture(new FileInfo(Path.Combine(Path.GetTempPath(), "does-not-exist.png")), PictureAlignment.Left));

		Assert.AreEqual("before", ws.HeaderFooter.OddHeader.LeftAlignedText, "rejected data must not leave a picture placeholder behind");
	}

	[TestMethod]
	public void HeaderFooter_InsertPicture_NullArguments_AreRejected()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("HF");

		Assert.ThrowsException<ArgumentNullException>(() => ws.HeaderFooter.OddHeader.InsertPicture((byte[])null, PictureAlignment.Left));
		Assert.ThrowsException<ArgumentNullException>(() => ws.HeaderFooter.OddHeader.InsertPicture((Stream)null, PictureAlignment.Left));
		Assert.ThrowsException<ArgumentNullException>(() => ws.HeaderFooter.OddHeader.InsertPicture((FileInfo)null, PictureAlignment.Left));
	}

	[TestMethod]
	public void AddPicture_NullFileInfo_ThrowsArgumentNullException()
	{
		using var package = new ExcelPackage();
		var ws = package.Workbook.Worksheets.Add("Pictures");

		Assert.ThrowsException<ArgumentNullException>(() => ws.Drawings.AddPicture("P", (FileInfo)null, null));
	}
}

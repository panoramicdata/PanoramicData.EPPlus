using OfficeOpenXml.Compatibility;
using System.Buffers.Binary;

namespace PanoramicData.EPPlus.Test;

public partial class ImageHeaderReaderTests
{
	[TestMethod]
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

		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(tiff));
	}

	[TestMethod]
	public void Tiff_WithoutWidthAndHeight_IsRejected()
	{
		var tiff = BuildTiff(true, new TiffEntry(282, Rational, 0, (72, 1)));

		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(tiff));
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

		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(valid[..10]), "cut inside the IFD");
		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(valid[..^4]), "cut inside the rational data");

		var wildIfd = (byte[])valid.Clone();
		BinaryPrimitives.WriteUInt32LittleEndian(wildIfd.AsSpan(4), 0xFFFFFFF0);
		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(wildIfd), "IFD offset beyond the data");

		var wildRational = (byte[])valid.Clone();
		BinaryPrimitives.WriteUInt32LittleEndian(wildRational.AsSpan(10 + (2 * 12) + 8), 0x7FFFFFF0);
		Assert.ThrowsExactly<InvalidDataException>(() => ImageHeaderReader.Read(wildRational), "rational offset beyond the data");
	}
}

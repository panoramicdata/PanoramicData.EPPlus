using System.Buffers.Binary;

namespace OfficeOpenXml.Compatibility;

/// <summary>
/// The pixel size, resolution and content type of an image, read from its file header.
/// </summary>
/// <remarks>
/// Pure managed code: no image is decoded and no graphics library is needed, so it behaves
/// identically on every platform. Handles PNG, JPEG, GIF, BMP, TIFF, EMF and WMF (placeable).
/// </remarks>
internal readonly record struct ImageHeaderInfo(string ContentType, int Width, int Height, float HorizontalResolution, float VerticalResolution)
{
	internal const float DefaultDpi = 96F;
}

internal static class ImageHeaderReader
{
	/// <summary>
	/// Sniffs the content type from the signature. Unrecognised data is reported as image/jpeg,
	/// matching the long-standing fallback.
	/// </summary>
	internal static string GetContentType(ReadOnlySpan<byte> b) => GetRasterContentType(b) ?? GetVectorContentType(b) ?? "image/jpeg";

	private static string GetRasterContentType(ReadOnlySpan<byte> b)
	{
		if (b.StartsWith((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47]))
		{
			return "image/png";
		}

		if (b.StartsWith((ReadOnlySpan<byte>)[0xFF, 0xD8, 0xFF]))
		{
			return "image/jpeg";
		}

		if (b.StartsWith((ReadOnlySpan<byte>)[0x47, 0x49, 0x46]))
		{
			return "image/gif";
		}

		if (b.StartsWith((ReadOnlySpan<byte>)[0x42, 0x4D]))
		{
			return "image/bmp";
		}

		return null;
	}

	private static string GetVectorContentType(ReadOnlySpan<byte> b)
	{
		if (b.StartsWith((ReadOnlySpan<byte>)[0x49, 0x49, 0x2A, 0]) || b.StartsWith((ReadOnlySpan<byte>)[0x4D, 0x4D, 0, 0x2A]))
		{
			return "image/x-tiff";
		}

		if (b.Length >= 44 && BinaryPrimitives.ReadInt32LittleEndian(b) == 1 && BinaryPrimitives.ReadInt32LittleEndian(b[40..]) == 0x464D4520)
		{
			return "image/x-emf";
		}

		if (b.StartsWith((ReadOnlySpan<byte>)[0xD7, 0xCD, 0xC6, 0x9A]))
		{
			return "image/x-wmf";
		}

		return null;
	}

	/// <summary>
	/// Reads the header. Throws <see cref="InvalidDataException"/> when the data is not a
	/// recognised image or is truncated.
	/// </summary>
	internal static ImageHeaderInfo Read(byte[] bytes)
	{
		ArgumentNullException.ThrowIfNull(bytes);
		try
		{
			var contentType = GetContentType(bytes);
			var data = new ReadOnlySpan<byte>(bytes);
			var (w, h, hx, hy) = contentType switch
			{
				"image/png" => ReadPng(data),
				"image/gif" => ReadGif(data),
				"image/bmp" => ReadBmp(data),
				"image/x-tiff" => ReadTiff(data),
				"image/x-emf" => ReadEmf(data),
				"image/x-wmf" => ReadWmf(data),
				_ => ReadJpeg(data),
			};
			return new ImageHeaderInfo(
				contentType,
				w,
				h,
				hx > 0 ? hx : ImageHeaderInfo.DefaultDpi,
				hy > 0 ? hy : ImageHeaderInfo.DefaultDpi);
		}
		catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException or OverflowException)
		{
			throw new InvalidDataException("Data is not a supported image or is corrupt", ex);
		}
	}

	private static (int, int, float, float) ReadPng(ReadOnlySpan<byte> b)
	{
		// 8 byte signature, then chunks of length(4) type(4) data crc(4). IHDR is first.
		var width = BinaryPrimitives.ReadInt32BigEndian(b[16..]);
		var height = BinaryPrimitives.ReadInt32BigEndian(b[20..]);
		float hx = 0, hy = 0;
		var pos = 8;
		while (pos + 12 <= b.Length)
		{
			var len = BinaryPrimitives.ReadInt32BigEndian(b[pos..]);
			var type = b.Slice(pos + 4, 4);
			if (type.SequenceEqual("pHYs"u8) && len >= 9 && b[pos + 16] == 1)
			{
				hx = (float)(BinaryPrimitives.ReadUInt32BigEndian(b[(pos + 8)..]) * 0.0254);
				hy = (float)(BinaryPrimitives.ReadUInt32BigEndian(b[(pos + 12)..]) * 0.0254);
				break;
			}

			if (type.SequenceEqual("IDAT"u8) || type.SequenceEqual("IEND"u8) || len < 0)
			{
				break; // pHYs must precede IDAT
			}

			pos += 12 + len;
		}

		return (width, height, hx, hy);
	}

	private static (int, int, float, float) ReadGif(ReadOnlySpan<byte> b)
		=> (BinaryPrimitives.ReadUInt16LittleEndian(b[6..]), BinaryPrimitives.ReadUInt16LittleEndian(b[8..]), 0, 0);

	private static (int, int, float, float) ReadBmp(ReadOnlySpan<byte> b)
	{
		var headerSize = BinaryPrimitives.ReadInt32LittleEndian(b[14..]);
		if (headerSize == 12)
		{
			// BITMAPCOREHEADER
			return (BinaryPrimitives.ReadUInt16LittleEndian(b[18..]), BinaryPrimitives.ReadUInt16LittleEndian(b[20..]), 0, 0);
		}

		var width = BinaryPrimitives.ReadInt32LittleEndian(b[18..]);
		var height = Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(b[22..])); // negative = top-down
		var xppm = BinaryPrimitives.ReadInt32LittleEndian(b[38..]);
		var yppm = BinaryPrimitives.ReadInt32LittleEndian(b[42..]);
		return (width, height, (float)(xppm * 0.0254), (float)(yppm * 0.0254));
	}

	private static (int, int, float, float) ReadJpeg(ReadOnlySpan<byte> b)
	{
		float hx = 0, hy = 0;
		var pos = 2;
		while (pos + 4 <= b.Length)
		{
			if (b[pos] != 0xFF)
			{
				pos++;
				continue;
			}

			var marker = b[pos + 1];
			if (marker == 0xFF)
			{
				pos++; // fill byte
				continue;
			}

			if (marker is 0xD8 or 0x01 or (>= 0xD0 and <= 0xD7))
			{
				pos += 2; // standalone markers have no length
				continue;
			}

			var len = BinaryPrimitives.ReadUInt16BigEndian(b[(pos + 2)..]);
			ReadJfifDensity(b, pos, marker, len, ref hx, ref hy);

			// Start-of-frame markers (not DHT 0xC4, JPG 0xC8, DAC 0xCC)
			if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
			{
				var height = BinaryPrimitives.ReadUInt16BigEndian(b[(pos + 5)..]);
				var width = BinaryPrimitives.ReadUInt16BigEndian(b[(pos + 7)..]);
				return (width, height, hx, hy);
			}

			pos += 2 + len;
		}

		throw new InvalidDataException("No JPEG frame header found");
	}

	private static void ReadJfifDensity(ReadOnlySpan<byte> b, int pos, byte marker, ushort len, ref float hx, ref float hy)
	{
		if (marker != 0xE0 || len < 16 || !b.Slice(pos + 4, 5).SequenceEqual("JFIF\0"u8))
		{
			return;
		}

		var units = b[pos + 11];
		var xd = BinaryPrimitives.ReadUInt16BigEndian(b[(pos + 12)..]);
		var yd = BinaryPrimitives.ReadUInt16BigEndian(b[(pos + 14)..]);
		if (units == 1)
		{
			hx = xd;
			hy = yd;
		}
		else if (units == 2)
		{
			hx = (float)(xd * 2.54);
			hy = (float)(yd * 2.54);
		}
	}

	private static ushort TiffU16(ReadOnlySpan<byte> s, int o, bool little)
		=> little ? BinaryPrimitives.ReadUInt16LittleEndian(s[o..]) : BinaryPrimitives.ReadUInt16BigEndian(s[o..]);

	private static uint TiffU32(ReadOnlySpan<byte> s, int o, bool little)
		=> little ? BinaryPrimitives.ReadUInt32LittleEndian(s[o..]) : BinaryPrimitives.ReadUInt32BigEndian(s[o..]);

	// An unsupported field type is a hard error for the fields the size depends on, and is
	// ignored (0) for the optional resolution fields, which then fall back to the default DPI.
	private static double TiffValue(ReadOnlySpan<byte> s, int entry, bool little, bool required)
	{
		var type = TiffU16(s, entry + 2, little);
		switch (type)
		{
			case 3:
				return TiffU16(s, entry + 8, little);   // SHORT
			case 4:
				return TiffU32(s, entry + 8, little);   // LONG
			case 5:                                     // RATIONAL
				var at = (int)TiffU32(s, entry + 8, little);
				return (double)TiffU32(s, at, little) / Math.Max(1U, TiffU32(s, at + 4, little));
			default:
				return required
					? throw new InvalidDataException($"Unsupported TIFF field type {type} for a required field")
					: 0;
		}
	}

	private static (double Width, double Height, double XRes, double YRes, int Unit) ReadTiffTags(ReadOnlySpan<byte> b, int ifd, bool little)
	{
		var count = TiffU16(b, ifd, little);
		double width = 0, height = 0, xres = 0, yres = 0;
		var unit = 2;
		for (var i = 0; i < count; i++)
		{
			var entry = ifd + 2 + (i * 12);
			switch (TiffU16(b, entry, little))
			{
				case 256:
					width = TiffValue(b, entry, little, required: true);
					break;
				case 257:
					height = TiffValue(b, entry, little, required: true);
					break;
				case 282:
					xres = TiffValue(b, entry, little, required: false);
					break;
				case 283:
					yres = TiffValue(b, entry, little, required: false);
					break;
				case 296:
					unit = (int)TiffValue(b, entry, little, required: false);
					break;
				default:
					break; // every other tag is irrelevant to the size and resolution
			}
		}

		return (width, height, xres, yres, unit);
	}

	private static (int, int, float, float) ReadTiff(ReadOnlySpan<byte> b)
	{
		var little = b[0] == 0x49;
		var ifd = (int)TiffU32(b, 4, little);
		var (width, height, xres, yres, unit) = ReadTiffTags(b, ifd, little);

		if (width < 1 || height < 1)
		{
			throw new InvalidDataException("TIFF has no image width and height");
		}

		var scale = unit == 3 ? 2.54 : 1;
		return ((int)width, (int)height, (float)(xres * scale), (float)(yres * scale));
	}

	private static (int, int, float, float) ReadEmf(ReadOnlySpan<byte> b)
	{
		// ENHMETAHEADER: rclFrame (0.01mm units) at offset 24.
		var frameWidth = BinaryPrimitives.ReadInt32LittleEndian(b[32..]) - BinaryPrimitives.ReadInt32LittleEndian(b[24..]);
		var frameHeight = BinaryPrimitives.ReadInt32LittleEndian(b[36..]) - BinaryPrimitives.ReadInt32LittleEndian(b[28..]);
		var dpi = ImageHeaderInfo.DefaultDpi;
		return (
			(int)Math.Round(frameWidth / 2540.0 * dpi, MidpointRounding.AwayFromZero),
			(int)Math.Round(frameHeight / 2540.0 * dpi, MidpointRounding.AwayFromZero),
			dpi,
			dpi);
	}

	private static (int, int, float, float) ReadWmf(ReadOnlySpan<byte> b)
	{
		// Placeable header: bounding box in logical units, then units-per-inch.
		var left = BinaryPrimitives.ReadInt16LittleEndian(b[6..]);
		var top = BinaryPrimitives.ReadInt16LittleEndian(b[8..]);
		var right = BinaryPrimitives.ReadInt16LittleEndian(b[10..]);
		var bottom = BinaryPrimitives.ReadInt16LittleEndian(b[12..]);
		var inch = (int)BinaryPrimitives.ReadUInt16LittleEndian(b[14..]);
		if (inch == 0)
		{
			inch = 1440;
		}

		var dpi = ImageHeaderInfo.DefaultDpi;
		return (
			(int)Math.Round((right - left) / (double)inch * dpi, MidpointRounding.AwayFromZero),
			(int)Math.Round((bottom - top) / (double)inch * dpi, MidpointRounding.AwayFromZero),
			dpi,
			dpi);
	}
}

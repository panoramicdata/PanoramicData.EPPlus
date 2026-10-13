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
	internal static string GetContentType(ReadOnlySpan<byte> b)
	{
		if (b.Length >= 4 && b[0] == 0x89 && b[1] == 0x50 && b[2] == 0x4E && b[3] == 0x47)
		{
			return "image/png";
		}

		if (b.Length >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF)
		{
			return "image/jpeg";
		}

		if (b.Length >= 3 && b[0] == 0x47 && b[1] == 0x49 && b[2] == 0x46)
		{
			return "image/gif";
		}

		if (b.Length >= 2 && b[0] == 0x42 && b[1] == 0x4D)
		{
			return "image/bmp";
		}

		if (b.Length >= 4 && ((b[0] == 0x49 && b[1] == 0x49 && b[2] == 0x2A && b[3] == 0) || (b[0] == 0x4D && b[1] == 0x4D && b[2] == 0 && b[3] == 0x2A)))
		{
			return "image/x-tiff";
		}

		if (b.Length >= 44 && BinaryPrimitives.ReadInt32LittleEndian(b) == 1 && BinaryPrimitives.ReadInt32LittleEndian(b[40..]) == 0x464D4520)
		{
			return "image/x-emf";
		}

		if (b.Length >= 4 && b[0] == 0xD7 && b[1] == 0xCD && b[2] == 0xC6 && b[3] == 0x9A)
		{
			return "image/x-wmf";
		}

		return "image/jpeg";
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
			if (marker == 0xE0 && len >= 16 && b.Slice(pos + 4, 5).SequenceEqual("JFIF\0"u8))
			{
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

	private static (int, int, float, float) ReadTiff(ReadOnlySpan<byte> b)
	{
		var little = b[0] == 0x49;
		ushort U16(ReadOnlySpan<byte> s, int o) => little ? BinaryPrimitives.ReadUInt16LittleEndian(s[o..]) : BinaryPrimitives.ReadUInt16BigEndian(s[o..]);
		uint U32(ReadOnlySpan<byte> s, int o) => little ? BinaryPrimitives.ReadUInt32LittleEndian(s[o..]) : BinaryPrimitives.ReadUInt32BigEndian(s[o..]);
		// An unsupported field type is a hard error for the fields the size depends on, and is
		// ignored (0) for the optional resolution fields, which then fall back to the default DPI.
		double Value(ReadOnlySpan<byte> s, int entry, bool required)
		{
			var type = U16(s, entry + 2);
			switch (type)
			{
				case 3:
					return U16(s, entry + 8);   // SHORT
				case 4:
					return U32(s, entry + 8);   // LONG
				case 5:                         // RATIONAL
					var at = (int)U32(s, entry + 8);
					return (double)U32(s, at) / Math.Max(1U, U32(s, at + 4));
				default:
					return required
						? throw new InvalidDataException($"Unsupported TIFF field type {type} for a required field")
						: 0;
			}
		}

		var ifd = (int)U32(b, 4);
		var count = U16(b, ifd);
		double width = 0, height = 0, xres = 0, yres = 0;
		var unit = 2;
		for (var i = 0; i < count; i++)
		{
			var entry = ifd + 2 + (i * 12);
			switch (U16(b, entry))
			{
				case 256:
					width = Value(b, entry, required: true);
					break;
				case 257:
					height = Value(b, entry, required: true);
					break;
				case 282:
					xres = Value(b, entry, required: false);
					break;
				case 283:
					yres = Value(b, entry, required: false);
					break;
				case 296:
					unit = (int)Value(b, entry, required: false);
					break;
				default:
					break; // every other tag is irrelevant to the size and resolution
			}
		}

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

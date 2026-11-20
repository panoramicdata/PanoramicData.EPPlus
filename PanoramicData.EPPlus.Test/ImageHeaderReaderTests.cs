using System.Buffers.Binary;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// The managed image header reader, which replaces GDI+ and must behave identically on every platform:
/// TIFF, EMF and WMF dimensions and resolution, malformed and truncated data, and the header/footer
/// InsertPicture overloads that sit on top of it.
/// </summary>
[TestClass]
public partial class ImageHeaderReaderTests
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
}

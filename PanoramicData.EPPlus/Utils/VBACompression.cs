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
 *******************************************************************************
 * Jan Källman		Added		01-01-2012
 * Jan Källman      Added compression support 27-03-2012
 *******************************************************************************/
using System;
using System.IO;

namespace OfficeOpenXml.Utils;

internal static class VBACompression
{
	#region  Compression
	/// <summary>
	/// Compression using a run length encoding algorithm.
	/// See MS-OVBA Section 2.4
	/// </summary>
	/// <param name="part">Byte array to decompress</param>
	/// <returns></returns>
	internal static byte[] CompressPart(byte[] part)
	{
		MemoryStream ms = new(4096);
		BinaryWriter br = new(ms);
		br.Write((byte)1);

		var compStart = 1;
		var compEnd = 4098;
		var decompStart = 0;
		var decompEnd = part.Length < 4096 ? part.Length : 4096;

		while (decompStart < decompEnd && compStart < compEnd)
		{
			var chunk = CompressChunk(part, ref decompStart);
			if (chunk != null && chunk.Length > 0)
			{
				var header = (ushort)((chunk.Length - 1) & 0xFFF);
				header |= 0xB000;   //B=011 A=1
				br.Write(header);
				br.Write(chunk);
			}

			decompEnd = part.Length < decompStart + 4096 ? part.Length : decompStart + 4096;
		}


		br.Flush();
		return ms.ToArray();
	}
	private static byte[] CompressChunk(byte[] buffer, ref int startPos)
	{
		var comprBuffer = new byte[4096];
		var flagPos = 0;
		var cPos = 1;
		var dPos = startPos;
		var dEnd = startPos + 4096 < buffer.Length ? startPos + 4096 : buffer.Length;
		while (dPos < dEnd)
		{
			byte tokenFlags = 0;
			for (var i = 0; i < 8; i++)
			{
				if (CompressToken(buffer, startPos, dEnd, comprBuffer, ref dPos, ref cPos))
				{
					tokenFlags |= (byte)(1 << i);
				}

				if (dPos >= dEnd)
				{
					break;
				}
			}

			comprBuffer[flagPos] = tokenFlags;
			flagPos = cPos++;
		}

		var ret = new byte[cPos - 1];
		Array.Copy(comprBuffer, ret, ret.Length);
		startPos = dEnd;
		return ret;
	}

	// Returns true when a copy token was written, false for a literal byte.
	private static bool CompressToken(byte[] buffer, int startPos, int dEnd, byte[] comprBuffer, ref int dPos, ref int cPos)
	{
		if (dPos - startPos > 0)
		{
			var bitCount = GetLengthBits(dPos - startPos);
			var lengthMask = (ushort)(0xFFFF >> (16 - bitCount));
			var bestLength = FindBestMatch(buffer, startPos, dPos, dEnd, lengthMask, out var bestCandidate);
			if (bestLength >= 3)
			{
				var token = (ushort)((ushort)(dPos - (bestCandidate + 1)) << bitCount | (ushort)(bestLength - 3));
				Array.Copy(BitConverter.GetBytes(token), 0, comprBuffer, cPos, 2);
				dPos += bestLength;
				cPos += 2;
				return true;
			}
		}

		comprBuffer[cPos++] = buffer[dPos++];
		return false;
	}

	private static int FindBestMatch(byte[] buffer, int startPos, int dPos, int dEnd, ushort lengthMask, out int bestCandidate)
	{
		bestCandidate = -1;
		var bestLength = 0;
		for (var candidate = dPos - 1; candidate >= startPos; candidate--)
		{
			if (buffer[candidate] == buffer[dPos])
			{
				var length = GetMatchLength(buffer, candidate, dPos, dEnd, lengthMask);
				if (length > bestLength)
				{
					bestCandidate = candidate;
					bestLength = length;
					if (bestLength == lengthMask)
					{
						break;
					}
				}
			}
		}

		return bestLength;
	}

	private static int GetMatchLength(byte[] buffer, int candidate, int dPos, int dEnd, ushort lengthMask)
	{
		var length = 1;
		while (buffer.Length > dPos + length && buffer[candidate + length] == buffer[dPos + length] && length < lengthMask && dPos + length < dEnd)
		{
			length++;
		}

		return length;
	}
	internal static byte[] DecompressPart(byte[] part) => DecompressPart(part, 0);
	/// <summary>
	/// Decompression using a run length encoding algorithm.
	/// See MS-OVBA Section 2.4
	/// </summary>
	/// <param name="part">Byte array to decompress</param>
	/// <param name="startPos"></param>
	/// <returns></returns>
	internal static byte[] DecompressPart(byte[] part, int startPos)
	{

		if (part[startPos] != 1)
		{
			return null;
		}

		MemoryStream ms = new(4096);
		var compressPos = startPos + 1;
		while (compressPos < part.Length - 1)
		{
			DecompressChunk(ms, part, ref compressPos);
		}

		return ms.ToArray();
	}
	private static void DecompressChunk(MemoryStream ms, byte[] compBuffer, ref int pos)
	{
		var header = BitConverter.ToUInt16(compBuffer, pos);
		var size = (header & 0xFFF) + 3;
		var endPos = pos + size;
		var b = (header & 0x8000) >> 15;
		pos += 2;
		if (b == 1) //Compressed chunk
		{
			DecompressCompressedChunk(ms, compBuffer, ref pos, endPos);
		}
		else //Raw chunk
		{
			ms.Write(compBuffer, pos, size);
			pos += size;
		}
	}

	private static void DecompressCompressedChunk(MemoryStream ms, byte[] compBuffer, ref int pos, int endPos)
	{
		var decomprPos = 0;
		var buffer = new byte[4198]; //Add an extra 100 byte. Some workbooks have overflowing worksheets.
		while (pos < compBuffer.Length && pos < endPos)
		{
			//Decompress token
			var token = compBuffer[pos++];
			if (pos >= endPos)
			{
				break;
			}

			for (var i = 0; i < 8; i++)
			{
				//Literal token
				if ((token & (1 << i)) == 0)
				{
					ms.WriteByte(compBuffer[pos]);
					buffer[decomprPos++] = compBuffer[pos++];
				}
				else
				{
					buffer = DecompressCopyToken(ms, compBuffer, pos, buffer, ref decomprPos);
					pos += 2;
				}

				if (pos >= endPos)
				{
					break;
				}
			}
		}
	}

	private static byte[] DecompressCopyToken(MemoryStream ms, byte[] compBuffer, int pos, byte[] buffer, ref int decomprPos)
	{
		var t = BitConverter.ToUInt16(compBuffer, pos);
		var bitCount = GetLengthBits(decomprPos);
		var lengthMask = (ushort)(0xFFFF >> (16 - bitCount));
		var offsetMask = (ushort)~lengthMask;
		var length = (lengthMask & t) + 3;
		var offset = (offsetMask & t) >> bitCount;
		var source = decomprPos - offset - 1;
		if (decomprPos + length >= buffer.Length)
		{
			// Excel-generated VBA projects can overflow the nominal chunk size, so grow the buffer.
			var largerBuffer = new byte[buffer.Length + 4098];
			Array.Copy(buffer, largerBuffer, decomprPos);
			buffer = largerBuffer;
		}

		// Byte-wise, and into the buffer too: copy tokens can overlap bytes a later token reads.
		for (var c = 0; c < length; c++)
		{
			ms.WriteByte(buffer[source]);
			buffer[decomprPos++] = buffer[source++];
		}

		return buffer;
	}

	private static int GetLengthBits(int decompPos)
	{
		if (decompPos > 4096)
		{
			//We should never end up here, but if so this is the formula to calculate the bits...
			return 12 - (int)Math.Truncate(Math.Log(decompPos - 1 >> 4, 2) + 1);
		}

		var bits = 12;
		for (var limit = 16; decompPos > limit; limit <<= 1)
		{
			bits--;
		}

		return bits;
	}
	#endregion
}

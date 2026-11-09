using OfficeOpenXml.Utils;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	const int BLOCKSIZE = 8192;
	/// <summary>
	/// Extracts the workbook XML without the sheetData-element (containing all cell data).
	/// Xml-Cell data can be extreemly large (GB), so we find the sheetdata element in the streem (position start) and 
	/// then tries to find the &lt;/sheetData&gt; element from the end-parameter.
	/// This approach is to avoid out of memory exceptions reading large packages
	/// </summary>
	/// <param name="stream">the worksheet stream</param>
	/// <param name="start">Position from previous reading where we found the sheetData element</param>
	/// <param name="end">End position, where &lt;/sheetData&gt; or &lt;sheetData/&gt; is found</param>
	/// <param name="encoding">Encoding</param>
	/// <returns>The worksheet xml, with an empty sheetdata. (Sheetdata is in memory in the worksheet)</returns>
	private static string GetWorkSheetXml(Stream stream, long start, long end, out Encoding encoding)
	{
		StreamReader sr = new(stream);
		var length = 0;
		char[] block;
		int pos;
		StringBuilder sb = new();
		Match startmMatch, endMatch;
		do
		{
			var size = stream.Length < BLOCKSIZE ? (int)stream.Length : BLOCKSIZE;
			block = new char[size];
			pos = sr.ReadBlock(block, 0, size);
			sb.Append(block, 0, pos);
			length += size;
		}
		while (length < start + 20 && length < end);    //the  start-pos contains the stream position of the sheetData element. Add 20 (with some safty for whitespace, streampointer diff etc, just so be sure). 
		startmMatch = Regex.Match(sb.ToString(), string.Format("(<[^>]*{0}[^>]*>)", "sheetData"));
		if (!startmMatch.Success) //Not found
		{
			encoding = sr.CurrentEncoding;
			return sb.ToString();
		}
		else
		{
			var s = sb.ToString();
			var xml = s[..startmMatch.Index];
			if (ConvertUtil._invariantCompareInfo.IsSuffix(startmMatch.Value, "/>"))        //Empty sheetdata
			{
				xml += s[startmMatch.Index..];
			}
			else
			{
				if (sr.Peek() != -1)        //Now find the end tag </sheetdata> so we can add the end of the xml document
				{
					ReadSheetDataEndBlock(stream, end, ref sr, ref s);
				}

				endMatch = Regex.Match(s, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
				xml += "<sheetData/>" + s[(endMatch.Index + endMatch.Length)..];
			}

			if (sr.Peek() > -1)
			{
				xml += sr.ReadToEnd();
			}

			encoding = sr.CurrentEncoding;
			return xml;
		}
	}
	private static void ReadSheetDataEndBlock(Stream stream, long end, ref StreamReader sr, ref string s)
	{
		/**** Fixes issue 14788. Fix by Philip Garrett ****/
		var endSeekStart = end;

		while (endSeekStart >= 0)
		{
			endSeekStart = Math.Max(endSeekStart - BLOCKSIZE, 0);
			var size = (int)(end - endSeekStart);
			stream.Seek(endSeekStart, SeekOrigin.Begin);
			var block = new char[size];
			sr = new StreamReader(stream);
			var pos = sr.ReadBlock(block, 0, size);
			var sb = new StringBuilder();
			sb.Append(block, 0, pos);
			s = sb.ToString();
			var endMatch = Regex.Match(s, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
			if (endMatch.Success)
			{
				break;
			}
		}
	}

	private static void GetBlockPos(string xml, string tag, ref int start, ref int end)
	{
		Match startmMatch, endMatch;
		startmMatch = Regex.Match(xml[start..], string.Format("(<[^>]*{0}[^>]*>)", tag));

		if (!startmMatch.Success) //Not found
		{
			start = -1;
			end = -1;
			return;
		}

		var startPos = startmMatch.Index + start;
		if (startmMatch.Value.Substring(startmMatch.Value.Length - 2, 1) == "/")
		{
			end = startPos + startmMatch.Length;
		}
		else
		{
			endMatch = Regex.Match(xml[start..], string.Format("(</[^>]*{0}[^>]*>)", tag));
			if (endMatch.Success)
			{
				end = endMatch.Index + endMatch.Length + start;
			}
		}

		start = startPos;
	}
	private static bool ReadUntil(XmlReader xr, params string[] tagName)
	{
		if (xr.EOF)
		{
			return false;
		}

		while (!Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag)))
		{
			xr.Read();
			if (xr.EOF)
			{
				return false;
			}
		}

		return ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]);
	}
}

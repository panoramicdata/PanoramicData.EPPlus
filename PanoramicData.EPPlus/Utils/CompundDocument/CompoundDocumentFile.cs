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
 * Jan Källman		Added		28-MAR-2017
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
namespace OfficeOpenXml.Utils.CompundDocument;

/// <summary>
/// Read and write a compound document.
/// Read spec here https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/[MS-CFB].pdf
/// </summary>
internal partial class CompoundDocumentFile : IDisposable
{
	public CompoundDocumentFile()
	{
		RootItem = new CompoundDocumentItem { Name = "<Root>", Children = [], ObjectType = 5 };
		minorVersion = 0x3E;
		majorVersion = 3;
		const short sectorShift = 9;
		minSectorShift = 6;

		_sectorSize = 1 << sectorShift;
		_miniSectorSize = 1 << minSectorShift;
		_sectorSizeInt = _sectorSize / 4;
	}
	internal CompoundDocumentFile(FileInfo fi) : this(File.ReadAllBytes(fi.FullName))
	{

	}
	public CompoundDocumentFile(byte[] file) : this(new MemoryStream(file))
	{
	}
	public CompoundDocumentFile(MemoryStream ms)
	{
		ms.Seek(0, SeekOrigin.Begin);   //Fixes issue #60
		Read(new BinaryReader(ms));
	}
	private sealed class DocWriteInfo
	{
		internal List<int> DIFAT, FAT, miniFAT;
	}
	#region Constants
	const int miniFATSectorSize = 64;
	const int FATSectorSizeV3 = 512;
	const int FATSectorSizeV4 = 4096;

	const int DIFAT_SECTOR = -4; // 0xFFFFFFFC
	const int FAT_SECTOR = -3;   // 0xFFFFFFFD
	const int END_OF_CHAIN = -2; // 0xFFFFFFFE
	const int FREE_SECTOR = -1;  // 0xFFFFFFFF

	static readonly byte[] header = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
	#endregion
	#region Private Fields
	short minorVersion;
	short majorVersion;
	int numberOfDirectorySector;
	short minSectorShift;       //Bits for sector size

	int _numberOfFATSectors;            // (4 bytes): This integer field contains the count of the number of FAT sectors in the compound file.
	int _firstDirectorySectorLocation;  // (4 bytes): This integer field contains the starting sector number for the directory stream.            
	int _transactionSignatureNumber;    // (4 bytes): This integer field MAY contain a sequence number that is incremented every time the compound file is saved by an implementation that supports file transactions.This is the field that MUST be set to all zeroes if file transactions are not implemented.<1> 
	int _miniStreamCutoffSize;          // (4 bytes): This integer field MUST be set to 0x00001000. This field specifies the maximum size of a user-defined data stream that is allocated from the mini FAT and mini stream, and that cutoff is 4,096 bytes.Any user-defined data stream that is larger than or equal to this cutoff size must be allocated as normal sectors from the FAT.
	int _firstMiniFATSectorLocation;    // (4 bytes): This integer field contains the starting sector number for the mini FAT. 
	int _numberofMiniFATSectors;        // (4 bytes): This integer field contains the count of the number of mini FAT sectors in the compound file. 
	int _firstDIFATSectorLocation;      // (4 bytes): This integer field contains the starting sector number for the DIFAT. 
	int _numberofDIFATSectors;          // (4 bytes): This integer field contains the count of the number of DIFAT sectors in the compound file. 

	List<byte[]> _sectors, _miniSectors;
	int _sectorSize, _miniSectorSize;
	int _sectorSizeInt;
	int _currentDIFATSectorPos, _currentFATSectorPos, _currentDirSectorPos;
	int _prevDirFATSectorPos;

	#endregion
	public CompoundDocumentItem RootItem { get; set; }
	/// <summary>
	/// Verifies that the header is correct.
	/// </summary>
	/// <param name="fi">The file</param>
	/// <returns></returns>
	public static bool IsCompoundDocument(FileInfo fi)
	{
		try
		{
			using var fs = fi.OpenRead();
			var b = new byte[8];
			fs.ReadAtLeast(b.AsSpan(0, 8), 8, throwOnEndOfStream: false);
			return IsCompoundDocument(b);
		}
		catch
		{
			return false;
		}
	}
	public static bool IsCompoundDocument(MemoryStream ms)
	{
		var pos = ms.Position;
		ms.Position = 0;
		var b = new byte[8];
		ms.ReadAtLeast(b.AsSpan(0, 8), 8, throwOnEndOfStream: false);
		ms.Position = pos;
		return IsCompoundDocument(b);
	}
	public static bool IsCompoundDocument(byte[] b)
	{
		if (b == null || b.Length < 8)
		{
			return false;
		}

		for (var i = 0; i < 8; i++)
		{
			if (b[i] != header[i])
			{
				return false;
			}
		}

		return true;
	}
	#region Read
	internal void Read(BinaryReader br)
	{
		ReadHeader(br);
		var dwi = new DocWriteInfo { DIFAT = [], FAT = [], miniFAT = [] };

		for (var i = 0; i < 109; i++)
		{
			var d = br.ReadInt32();
			if (d >= 0)
			{
				dwi.DIFAT.Add(d);
			}
		}

		LoadSectors(br);
		if (_firstDIFATSectorLocation > 0)
		{
			LoadDIFATSectors(dwi);
		}

		dwi.FAT = ReadFAT(_sectors, dwi);
		var dir = ReadDirectories(_sectors, dwi);

		LoadMinSectors(dwi, dir);
		LoadStreams(dir, dwi);

		AddChildTree(dir[0], dir);
	}

	private void LoadStreams(List<CompoundDocumentItem> dir, DocWriteInfo dwi)
	{
		foreach (var d in dir)
		{
			if (d.Stream == null && d.StreamSize > 0)
			{
				if (d.StreamSize < _miniStreamCutoffSize)
				{
					d.Stream = GetStream(d.StartingSectorLocation, d.StreamSize, dwi.miniFAT, _miniSectors);
				}
				else
				{
					d.Stream = GetStream(d.StartingSectorLocation, d.StreamSize, dwi.FAT, _sectors);
				}
			}
		}
	}

	private void ReadHeader(BinaryReader br)
	{
		br.ReadBytes(8);    //Read header
		br.ReadBytes(16);   //Header CLSID (16 bytes): Reserved and unused class ID that MUST be set to all zeroes (CLSID_NULL).
		minorVersion = br.ReadInt16();
		majorVersion = br.ReadInt16();
		br.ReadInt16(); //Byte order
		var sectorShift = br.ReadInt16();
		minSectorShift = br.ReadInt16();

		_sectorSize = 1 << sectorShift;
		_miniSectorSize = 1 << minSectorShift;
		_sectorSizeInt = _sectorSize / 4;
		br.ReadBytes(6);    //Reserved
		numberOfDirectorySector = br.ReadInt32();
		_numberOfFATSectors = br.ReadInt32();
		_firstDirectorySectorLocation = br.ReadInt32();
		_transactionSignatureNumber = br.ReadInt32();
		_miniStreamCutoffSize = br.ReadInt32();
		_firstMiniFATSectorLocation = br.ReadInt32();
		_numberofMiniFATSectors = br.ReadInt32();
		_firstDIFATSectorLocation = br.ReadInt32();
		_numberofDIFATSectors = br.ReadInt32();
	}
	private void LoadDIFATSectors(DocWriteInfo dwi)
	{
		var nextSector = _firstDIFATSectorLocation;
		while (nextSector > 0)
		{
			var brDI = new BinaryReader(new MemoryStream(_sectors[nextSector]));
			var sect = -1;
			while (brDI.BaseStream.Position < _sectorSize)
			{
				if (sect > 0)
				{
					dwi.DIFAT.Add(sect);
				}

				sect = brDI.ReadInt32();
			}

			nextSector = sect;
		}
	}
	private void LoadSectors(BinaryReader br)
	{
		_sectors = [];
		while (br.BaseStream.Position < br.BaseStream.Length)
		{
			_sectors.Add(br.ReadBytes(_sectorSize));
		}
	}
	private void LoadMinSectors(DocWriteInfo dwi, List<CompoundDocumentItem> dir)
	{
		dwi.miniFAT = ReadMiniFAT(_sectors, dwi);
		dir[0].Stream = GetStream(dir[0].StartingSectorLocation, dir[0].StreamSize, dwi.FAT, _sectors);
		GetMiniSectors(dir[0].Stream);
	}
	private void GetMiniSectors(byte[] miniFATStream)
	{
		var br = new BinaryReader(new MemoryStream(miniFATStream));
		_miniSectors = [];
		while (br.BaseStream.Position < br.BaseStream.Length)
		{
			_miniSectors.Add(br.ReadBytes(_miniSectorSize));
		}
	}
	private static byte[] GetStream(int startingSectorLocation, long streamSize, List<int> FAT, List<byte[]> sectors)
	{
		var ms = new MemoryStream();
		var bw = new BinaryWriter(ms);

		var size = 0;
		var nextSector = startingSectorLocation;
		while (size < streamSize)
		{
			if (streamSize > size + sectors[nextSector].Length)
			{
				bw.Write(sectors[nextSector]);
				size += sectors[nextSector].Length;
			}
			else
			{
				var part = new byte[streamSize - size];
				Array.Copy(sectors[nextSector], part, (int)streamSize - size);
				bw.Write(part);
				size += part.Length;
			}

			nextSector = FAT[nextSector];
		}

		bw.Flush();
		return ms.ToArray();
	}
	private List<int> ReadMiniFAT(List<byte[]> sectors, DocWriteInfo dwi)
	{
		var l = new List<int>();
		var nextSector = _firstMiniFATSectorLocation;
		while (nextSector != END_OF_CHAIN)
		{
			var br = new BinaryReader(new MemoryStream(sectors[nextSector]));
			while (br.BaseStream.Position < _sectorSize)
			{
				var d = br.ReadInt32();
				l.Add(d);
			}

			nextSector = dwi.FAT[nextSector];
		}

		return l;
	}
	private List<CompoundDocumentItem> ReadDirectories(List<byte[]> sectors, DocWriteInfo dwi)
	{
		var dir = new List<CompoundDocumentItem>();
		var nextSector = _firstDirectorySectorLocation;
		while (nextSector != END_OF_CHAIN)
		{
			ReadDirectory(sectors, nextSector, dir);
			nextSector = dwi.FAT[nextSector];
		}

		return dir;
	}
	private List<int> ReadFAT(List<byte[]> sectors, DocWriteInfo dwi)
	{
		var l = new List<int>();
		foreach (var i in dwi.DIFAT)
		{
			var br = new BinaryReader(new MemoryStream(sectors[i]));
			while (br.BaseStream.Position < _sectorSize)
			{
				var d = br.ReadInt32();
				l.Add(d);
			}
		}

		return l;
	}
	private static void ReadDirectory(List<byte[]> sectors, int index, List<CompoundDocumentItem> l)
	{

		var br = new BinaryReader(new MemoryStream(sectors[index]));

		while (br.BaseStream.Position < br.BaseStream.Length)
		{
			var e = new CompoundDocumentItem();
			e.Read(br);
			if (e.ObjectType != 0)
			{
				l.Add(e);
			}
		}
	}
	internal void AddChildTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
	{
		if (e._handled == true)
		{
			return;
		}

		e._handled = true;
		if (e.ChildID > 0)
		{
			var c = dirs[e.ChildID];
			c.Parent = e;
			e.Children.Add(c);
			AddChildTree(c, dirs);
		}

		if (e.LeftSibling > 0)
		{
			var c = dirs[e.LeftSibling];
			c.Parent = e.Parent;
			c.Parent.Children.Insert(e.Parent.Children.IndexOf(e), c);
			AddChildTree(c, dirs);
		}

		if (e.RightSibling > 0)
		{
			var c = dirs[e.RightSibling];
			c.Parent = e.Parent;
			e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, c);
			AddChildTree(c, dirs);
		}

		if (e.ObjectType == 5)
		{
			RootItem = e;
		}
	}
	internal static void AddLeftSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
	{
		if (e.LeftSibling > 0)
		{
			var c = dirs[e.LeftSibling];
			if (c.Parent != null)
			{
				c.Parent = e.Parent;
				c.Parent.Children.Insert(e.Parent.Children.IndexOf(e), c);
				e._handled = true;
				AddLeftSiblingTree(c, dirs);
			}
		}
	}
	internal static void AddRightSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
	{
		if (e.RightSibling > 0)
		{
			var c = dirs[e.RightSibling];
			c.Parent = e.Parent;
			e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, c);
			e._handled = true;
			AddRightSiblingTree(c, dirs);
		}
	}
	#endregion
}


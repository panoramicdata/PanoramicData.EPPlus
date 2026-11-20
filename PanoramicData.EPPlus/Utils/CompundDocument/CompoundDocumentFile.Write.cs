namespace OfficeOpenXml.Utils.CompundDocument;

internal partial class CompoundDocumentFile
{
	#region Write
	public void Write(MemoryStream ms)
	{
		var bw = new BinaryWriter(ms);

		//InitValues
		minorVersion = 62;
		majorVersion = 3;
		minSectorShift = 6;             //2^6=64 bytes
		_miniStreamCutoffSize = 4096;
		_transactionSignatureNumber = 0;
		_firstDIFATSectorLocation = END_OF_CHAIN;
		_firstDirectorySectorLocation = 1;
		_firstMiniFATSectorLocation = 2;
		_numberOfFATSectors = 1;

		_currentDIFATSectorPos = 76;             //DIFAT Position in the header
		_currentFATSectorPos = _sectorSize;      //First FAT sector starts at Sector 0
		_currentDirSectorPos = _sectorSize * 2;  //First FAT sector starts at Sector 1
		_prevDirFATSectorPos = _sectorSize + 4;  //Dir sector starts FAT position 1 (4 for int size)

		bw.Write(new byte[512 * 4]);            //Allocate for Header and first FAT, Directory och MiniFAT sectors
		WritePosition(bw, 0, ref _currentDIFATSectorPos, false);
		WritePosition(bw, [FAT_SECTOR, END_OF_CHAIN, END_OF_CHAIN], ref _currentFATSectorPos);  //First sector is first FAT sector, second is First Dir sector, thirs is first Mini FAT sector.

		var dirs = FlattenDirs();

		//Write directories
		WriteDirs(bw, dirs);

		//Fill empty DISectors up to 109
		FillDIFAT(bw);
		//Finally write the header information in the top of the file
		WriteHeader(bw);
	}

	private List<CompoundDocumentItem> FlattenDirs()
	{
		var l = new List<CompoundDocumentItem>();
		InitItem(RootItem);
		l.Add(RootItem);
		RootItem.ChildID = AddChildren(RootItem, l);
		return l;
	}

	private static void InitItem(CompoundDocumentItem item)
	{
		item.LeftSibling = -1;
		item.RightSibling = -1;
		item._handled = false;
	}

	private int AddChildren(CompoundDocumentItem item, List<CompoundDocumentItem> l)
	{
		var childId = -1;
		item.ColorFlag = 1; //Always Black-No matter here, we just add nodes as a b-tree
		if (item.Children.Count > 0)
		{
			foreach (var c in item.Children)
			{
				InitItem(c);
			}

			item.Children.Sort();

			childId = SetSiblings(l.Count, item.Children, 0, item.Children.Count - 1, -1);
			l.AddRange(item.Children);
			foreach (var c in item.Children)
			{
				c.ChildID = AddChildren(c, l);
			}
		}

		return childId;
	}

	private static int SetSiblings(int listAdd, List<CompoundDocumentItem> children, int fromPos, int toPos, int currSibl)
	{
		var pos = GetPos(fromPos, toPos);

		var item = children[pos];
		if (item._handled)
		{
			return currSibl;
		}

		item._handled = true;
		if (fromPos == toPos)
		{
			return fromPos + listAdd;
		}

		var div = Math.Max(pos / 2, 1);

		var lPos = GetPos(fromPos, pos - 1);
		var rPos = GetPos(pos + 1, toPos);
		if (div == 1 && BothHandled(children, lPos, rPos))
		{
			return pos + listAdd;
		}

		if (lPos > -1 && lPos >= fromPos)
		{
			item.LeftSibling = SetSiblings(listAdd, children, fromPos, pos - 1, item.LeftSibling);
		}

		if (rPos < children.Count && rPos <= toPos)
		{
			item.RightSibling = SetSiblings(listAdd, children, pos + 1, toPos, item.RightSibling);
		}

		return pos + listAdd;
	}

	private static bool BothHandled(List<CompoundDocumentItem> children, int lPos, int rPos) => children[lPos]._handled && children[rPos]._handled;

	private static int GetPos(int fromPos, int toPos)
	{
		var div = (toPos - fromPos) / 2;
		return fromPos + div;
	}

	private void FillDIFAT(BinaryWriter bw)
	{
		if (_currentDIFATSectorPos < _sectorSize)
		{
			bw.Seek(_currentDIFATSectorPos, SeekOrigin.Begin);
			while (_currentDIFATSectorPos < _sectorSize)
			{
				if (_currentDIFATSectorPos < 512)
				{
					bw.Write(0xFFFFFFFF);
				}
				else
				{
					bw.Write(0x0);
				}

				_currentDIFATSectorPos += 4;
			}
		}
	}

	private void WritePosition(BinaryWriter bw, int sector, ref int writePos, bool isFATEntry)
	{
		var pos = (int)bw.BaseStream.Position;
		bw.Seek(writePos, SeekOrigin.Begin);
		bw.Write(sector);
		writePos += 4;
		if (isFATEntry)
		{
			CheckUpdateDIFAT(bw);
		}

		bw.Seek(pos, SeekOrigin.Begin);
	}
	private static void WritePosition(BinaryWriter bw, int[] sectors, ref int writePos)
	{
		var pos = (int)bw.BaseStream.Position;
		bw.Seek(writePos, SeekOrigin.Begin);
		foreach (var sector in sectors)
		{
			bw.Write(sector);
			writePos += 4;
		}

		bw.Seek(pos, SeekOrigin.Begin);
	}
	private void WriteDirs(BinaryWriter bw, List<CompoundDocumentItem> dirs)
	{
		var miniFAT = SetMiniStream(dirs);
		AllocateFAT(bw, miniFAT.Length, dirs);
		WriteMiniFAT(bw, miniFAT);
		foreach (var entity in dirs)
		{
			if (entity.ObjectType == 5 || entity.StreamSize > _miniStreamCutoffSize)
			{
				entity.StartingSectorLocation = WriteStream(bw, entity.Stream);
			}
		}

		WriteDirStream(bw, dirs);
	}

	private void WriteDirStream(BinaryWriter bw, List<CompoundDocumentItem> dirs)
	{
		if (dirs.Count == 0)
		{
			return;
		}

		//First dirtory sector goes into sector 2
		bw.Seek((_firstDirectorySectorLocation + 1) * _sectorSize, SeekOrigin.Begin);
		for (var i = 0; i < Math.Min(_sectorSize / 128, dirs.Count); i++)
		{
			dirs[i].Write(bw);
		}

		bw.Seek(0, SeekOrigin.End);
		var start = (int)bw.BaseStream.Position / _sectorSize - 1;
		var pos = _sectorSize + 4;
		WritePosition(bw, start, ref pos, false);
		var streamLength = 0;
		for (var i = 4; i < dirs.Count; i++)
		{
			dirs[i].Write(bw);
			streamLength += 128;
		}

		WriteStreamFullSector(bw, _sectorSize);
		WriteFAT(bw, start, streamLength);
	}

	private void WriteMiniFAT(BinaryWriter bw, byte[] miniFAT)
	{
		if (miniFAT.Length >= _sectorSize)
		{
			bw.Seek((_firstMiniFATSectorLocation + 1) * _sectorSize, SeekOrigin.Begin);
			bw.Write(miniFAT, 0, _sectorSize);
			bw.Seek(0, SeekOrigin.End);
			if (miniFAT.Length > _sectorSize)
			{
				//Write next minifat sector to fat for sector 2
				var sector = ((int)bw.BaseStream.Position / _sectorSize) - 1;
				var pos = _sectorSize + (4 * 2);
				WritePosition(bw, sector, ref pos, false);

				//Write overflowing FAT sectors
				var b = new byte[miniFAT.Length - _sectorSize];
				Array.Copy(miniFAT, _sectorSize, b, 0, b.Length);
				WriteStream(bw, b);
			}

			_numberofMiniFATSectors = (miniFAT.Length + 1) / _sectorSize;
		}
	}

	private int WriteStream(BinaryWriter bw, byte[] stream)
	{
		bw.Seek(0, SeekOrigin.End);
		var start = (int)bw.BaseStream.Position / _sectorSize - 1;
		bw.Write(stream);
		WriteStreamFullSector(bw, _sectorSize);
		WriteFAT(bw, start, stream.Length);
		return start;
	}
	private void WriteFAT(BinaryWriter bw, int sector, long size)
	{
		bw.Seek(_currentFATSectorPos, SeekOrigin.Begin);
		var pos = _sectorSize;
		while (size > pos)
		{
			bw.Write(++sector);
			pos += _sectorSize;
			CheckUpdateDIFAT(bw);
		}

		bw.Write(END_OF_CHAIN);
		CheckUpdateDIFAT(bw);
		_currentFATSectorPos = (int)bw.BaseStream.Position;
		bw.Seek(0, SeekOrigin.End);
	}

	private void CheckUpdateDIFAT(BinaryWriter bw)
	{
		if (bw.BaseStream.Position % _sectorSize == 0)
		{
			if (_currentDIFATSectorPos % _sectorSize == 0)
			{
				bw.Seek(512, SeekOrigin.Current);
			}
			else if (bw.BaseStream.Position == (_sectorSize * 2))
			{
				bw.Seek(4 * _sectorSize, SeekOrigin.Begin);    //FAT continues after initizal dir och minifat sectors.
			}
			//Add to DIFAT
			var FATSector = (int)(bw.BaseStream.Position / _sectorSize - 1);
			WritePosition(bw, FATSector, ref _currentDIFATSectorPos, false);
			_numberOfFATSectors++;
			if (_currentDIFATSectorPos == _sectorSize || ((_currentDIFATSectorPos + 4) % _sectorSize == 0 && _currentDIFATSectorPos > _sectorSize))
			{
				AddDIFATSector(bw, FATSector);
			}
		}
	}

	private void AddDIFATSector(BinaryWriter bw, int FATSector)
	{
		bw.Write(new byte[_sectorSize]); //Write pre FAT sector
		if (_currentDIFATSectorPos > _sectorSize)                       //Write link to next DIFAT sector
		{
			WritePosition(bw, FATSector + 1, ref _currentDIFATSectorPos, false);
		}
		else
		{
			_firstDIFATSectorLocation = FATSector + 1;                    //Current sector
		}

		_currentDIFATSectorPos = (int)bw.BaseStream.Position;
		//Fill sector
		for (var i = 0; i < _sectorSize; i++)
		{
			bw.Write((byte)0xFF);
		}

		bw.Seek(-(_sectorSize * 2), SeekOrigin.Current);
	}

	private long GetFullStreamSize(int miniFatLength, List<CompoundDocumentItem> dirs)
	{
		var fullStreamSize = (long)miniFatLength - _sectorSize; //MiniFAT starts from sector 2, by default.
		foreach (var entity in dirs)
		{
			if (entity.ObjectType == 5 || entity.StreamSize > _miniStreamCutoffSize)
			{
				var rest = _sectorSize - entity.StreamSize % _sectorSize;
				fullStreamSize += entity.StreamSize;
				if (rest > 0 && rest < _sectorSize)
				{
					fullStreamSize += rest;
				}
			}
		}

		return fullStreamSize;
	}

	private void AllocateFAT(BinaryWriter bw, int miniFatLength, List<CompoundDocumentItem> dirs)
	{
		/*** First calculate full size ***/
		var noOfSectors = GetFullStreamSize(miniFatLength, dirs) / _sectorSize;

		//Directory Size
		var dirsPerSector = _sectorSize / 128;
		if (dirs.Count > dirsPerSector)
		{
			var dirSectors = GetSectors(dirs.Count, dirsPerSector);
			noOfSectors += dirSectors - 1; //Four item per sector. Sector two is fixed for directories
		}

		//First calc fat no sectors and difat sectors from full size
		var numberOfFATSectors = GetSectors((int)noOfSectors, _sectorSizeInt);       //Sector 0 is already allocated
		_numberofDIFATSectors = GetDIFatSectors(numberOfFATSectors);
		noOfSectors += numberOfFATSectors + _numberofDIFATSectors;

		//Calc fat sectors again with the added fat and di fat sectors.
		numberOfFATSectors = GetSectors((int)noOfSectors, _sectorSizeInt) + _numberofDIFATSectors;
		_numberofDIFATSectors = GetDIFatSectors(numberOfFATSectors);

		//Allocate FAT and DIFAT Sectors
		bw.Write(new byte[(numberOfFATSectors + (_numberofDIFATSectors > 0 ? _numberofDIFATSectors - 1 : 0)) * _sectorSize]);

		//Move to FAT Second sector (4).
		bw.Seek(_currentFATSectorPos, SeekOrigin.Begin);
		WriteFATAndDIFATEntries(bw, numberOfFATSectors);
		bw.Seek(0, SeekOrigin.End);
	}

	private void WriteFATAndDIFATEntries(BinaryWriter bw, int numberOfFATSectors)
	{
		var sectorPos = 1;
		for (var i = 1; i < 109; i++)     //We have 1 FAT sector to start with at sector 0
		{
			if (i < numberOfFATSectors + _numberofDIFATSectors)
			{
				WriteFATItem(bw, FAT_SECTOR);
				sectorPos++;
			}
			else
			{
				WriteFATItem(bw, END_OF_CHAIN);
				break;
			}
		}

		if (_numberofDIFATSectors > 0)
		{
			_firstDIFATSectorLocation = sectorPos + 1;
		}

		for (var j = 0; j < _numberofDIFATSectors; j++)
		{
			WriteFATItem(bw, DIFAT_SECTOR);
			for (var i = 0; i < _sectorSizeInt - 1; i++)
			{
				WriteFATItem(bw, FAT_SECTOR);
				sectorPos++;
				if (sectorPos >= numberOfFATSectors)
				{
					break;
				}
			}

			if (sectorPos > numberOfFATSectors)
			{
				break;
			}
		}
	}

	private int GetDIFatSectors(int FATSectors) => FATSectors > 109 ? GetSectors((FATSectors - 109), _sectorSizeInt - 1) : 0;

	private void WriteFATItem(BinaryWriter bw, int value)
	{
		bw.Write(value);
		CheckUpdateDIFAT(bw);
		_currentFATSectorPos = (int)bw.BaseStream.Position;
	}

	private static int GetSectors(int v, int size) => v % size == 0 ? v / size : v / size + 1;

	private byte[] SetMiniStream(List<CompoundDocumentItem> dirs)
	{
		//Create the miniStream
		var ms = new MemoryStream();
		var bwMiniFATStream = new BinaryWriter(ms);
		var bwMiniFAT = new BinaryWriter(new MemoryStream());
		var pos = 0;
		foreach (var entity in dirs)
		{
			if (entity.ObjectType != 5 && entity.StreamSize > 0 && entity.StreamSize <= _miniStreamCutoffSize)
			{
				bwMiniFATStream.Write(entity.Stream);
				WriteStreamFullSector(bwMiniFATStream, miniFATSectorSize);
				var size = _miniSectorSize;
				entity.StartingSectorLocation = pos;
				while (entity.StreamSize > size)
				{
					bwMiniFAT.Write(++pos);
					size += _miniSectorSize;
				}

				bwMiniFAT.Write(END_OF_CHAIN);
				pos++;
			}
		}

		dirs[0].StreamSize = ms.Length;
		dirs[0].Stream = ms.ToArray();

		WriteStreamFullSector(bwMiniFAT, _sectorSize);
		return ((MemoryStream)bwMiniFAT.BaseStream).ToArray();
	}

	private static void WriteStreamFullSector(BinaryWriter bw, int sectorSize)
	{
		var rest = sectorSize - (bw.BaseStream.Length % sectorSize);
		if (rest > 0 && rest < sectorSize)
		{
			bw.Write(new byte[rest]);
		}
	}
	private void WriteHeader(BinaryWriter bw)
	{
		bw.Seek(0, SeekOrigin.Begin);
		bw.Write(header);
		bw.Write(new byte[16]);             //ClsID all zero's
		bw.Write((short)0x3E);              //This field SHOULD be set to 0x003E if the major version field is either 0x0003 or 0x0004.                 
		bw.Write((short)0x3);               //Version 3
		bw.Write((ushort)0xFFFE);           // This field MUST be set to 0xFFFE. This field is a byte order mark for all integer fields, specifying little-endian byte order. 
		bw.Write((short)9);                 //Sector Shift
		bw.Write((short)6);                 //Mini Sector Shift
		bw.Write(new byte[6]);              //reserved
		bw.Write(0);                        //Number of Directory Sectors, unsupported i v3. Set to zero
		bw.Write(_numberOfFATSectors);      //Number of FAT Sectors
		bw.Write(1);                        //First Directory Sector Location
		bw.Write(0);                        //Transaction Signature Number
		bw.Write(_miniStreamCutoffSize);     //Mini Stream Cutoff Size
		bw.Write(2);                        //First Mini FAT Sector Location
		bw.Write(_numberofMiniFATSectors);   //Number of MiniFAT sectors
		bw.Write(_firstDIFATSectorLocation); //First DIFAT Sector Location
		bw.Write(_numberofDIFATSectors);     //Number of DIFAT Sectors
	}

	public void Dispose()
	{
		_miniSectors = null;
		_sectors = null;
	}
	#endregion
}

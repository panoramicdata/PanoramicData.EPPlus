using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Packaging;
using System;
using System.IO;
using System.IO.Compression;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// A workbook beyond the classic ZIP limits (4 GB, or 65,534 entries) must save with Zip64, and a workbook within them must not.
/// </summary>
[TestClass]
public class Zip64Tests
{
	private const int PartsBeyondTheLimit = 65_600;

	private static readonly byte[] _zip64EndOfCentralDirectory = [0x50, 0x4B, 0x06, 0x06];

	[TestMethod]
	public void ZipPackage_MoreThan65534Parts_SavesAndReadsBack()
	{
		var bytes = SavePackage(PartsBeyondTheLimit);

		using (var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read))
		{
			Assert.AreEqual(PartsBeyondTheLimit + 2, archive.Entries.Count);
			Assert.IsNotNull(archive.GetEntry("part/0.txt"));
			Assert.IsNotNull(archive.GetEntry($"part/{PartsBeyondTheLimit - 1}.txt"));
		}

		var reread = new ZipPackage(new MemoryStream(bytes));
		Assert.IsTrue(reread.PartExists(new Uri($"/part/{PartsBeyondTheLimit - 1}.txt", UriKind.Relative)));
	}

	[TestMethod]
	public void ZipPackage_MoreThan65534Parts_WritesZip64Records()
		=> Assert.IsTrue(Contains(SavePackage(PartsBeyondTheLimit), _zip64EndOfCentralDirectory));

	[TestMethod]
	public void ZipPackage_SmallPackage_WritesNoZip64Records()
		=> Assert.IsFalse(Contains(SavePackage(3), _zip64EndOfCentralDirectory));

	private static byte[] SavePackage(int partCount)
	{
		var package = new ZipPackage();
		for (var i = 0; i < partCount; i++)
		{
			var part = package.CreatePart(new Uri($"/part/{i}.txt", UriKind.Relative), "text/plain");
			part.GetStream().WriteByte((byte)'x');
		}

		using var stream = new MemoryStream();
		package.Save(stream);
		return stream.ToArray();
	}

	private static bool Contains(byte[] haystack, byte[] needle)
		=> haystack.AsSpan().IndexOf(needle) >= 0;
}

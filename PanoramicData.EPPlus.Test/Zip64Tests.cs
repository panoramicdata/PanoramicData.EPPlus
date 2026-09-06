using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Packaging.DotNetZip;
using OfficeOpenXml.Packaging.Ionic.Zip;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// MagicSuite: a LogicMonitor Health Check against a 25,663-resource portal failed after eleven
/// hours with
///
///   ZipException: Compressed or Uncompressed size, or offset exceeds the maximum value.
///                 Consider setting the UseZip64WhenSaving property on the ZipFile instance.
///     at ZipEntry.SetZip64Flags()
///     at ZipOutputStream.PutNextEntry(String)
///     at ZipPackage.Save(Stream)
///     at ExcelPackage.Save()
///
/// The workbook's content had been built successfully; only the save failed, so eleven hours of
/// report was lost at the final step.
///
/// Cause: ZipPackage.Save creates its ZipOutputStream setting only CompressionLevel, so
/// EnableZip64 keeps the constructor default of Zip64Option.Never (ZipOutputStream.cs). That caps
/// every workbook EPPlus writes at the classic ZIP limits: 4 GB of size or offset, and 65,534
/// entries.
/// </summary>
[TestClass]
public class Zip64Tests
{
	private const int EntriesBeyondTheLimit = 65_600;

	/// <summary>
	/// These tests drive ZipOutputStream directly rather than through ExcelPackage, so the
	/// code-page provider that IBM437 relies on has not already been registered for us.
	/// </summary>
	[ClassInitialize]
	public static void ClassInitialize(TestContext _)
		=> Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

	/// <summary>
	/// The 4 GB limit cannot be exercised in a unit test in reasonable time, but it and the entry
	/// limit share one cause: the Zip64Option.Never default. Pinning that default keeps the tests
	/// below demonstrably about the real defect.
	/// </summary>
	[TestMethod]
	public void ZipOutputStream_Default_IsZip64Never()
	{
		using var stream = new MemoryStream();
		using var zip = new ZipOutputStream(stream, leaveOpen: true);

		Assert.AreEqual(Zip64Option.Never, zip.EnableZip64);
	}

	/// <summary>
	/// Reproduces the production failure through the entry-count limit, which is reachable quickly.
	/// </summary>
	[TestMethod]
	public void ZipOutputStream_Zip64Never_FailsOnceEntryCountExceeds65534()
	{
		using var stream = new MemoryStream();

		var exception = Assert.ThrowsException<InvalidOperationException>(() =>
		{
			using var zip = new ZipOutputStream(stream, leaveOpen: true)
			{
				EnableZip64 = Zip64Option.Never
			};
			WriteEntries(zip, EntriesBeyondTheLimit);
		});

		StringAssert.Contains(exception.Message, "EnableZip64");
	}

	/// <summary>
	/// The proposed fix: the same workload succeeds, and the result is a readable archive holding
	/// every entry rather than merely a stream of bytes that did not throw.
	/// </summary>
	[TestMethod]
	public void ZipOutputStream_Zip64AsNecessary_WritesAndReadsBackMoreThan65534Entries()
	{
		using var stream = new MemoryStream();

		using (var zip = new ZipOutputStream(stream, leaveOpen: true)
		{
			EnableZip64 = Zip64Option.AsNecessary
		})
		{
			WriteEntries(zip, EntriesBeyondTheLimit);
		}

		stream.Position = 0;
		using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

		Assert.AreEqual(EntriesBeyondTheLimit, archive.Entries.Count);
		Assert.IsNotNull(archive.GetEntry("entry-0.txt"));
		Assert.IsNotNull(archive.GetEntry($"entry-{EntriesBeyondTheLimit - 1}.txt"));
	}

	/// <summary>
	/// AsNecessary rather than Always matters, so this checks the no-regression claim. The bytes
	/// are not identical, because the central directory differs, but a small archive must remain
	/// readable by a standard reader and hold exactly the same entries.
	/// </summary>
	[TestMethod]
	public void ZipOutputStream_Zip64AsNecessary_LeavesSmallArchivesReadable()
	{
		var never = ReadBackEntryNames(Zip64Option.Never);
		var asNecessary = ReadBackEntryNames(Zip64Option.AsNecessary);

		CollectionAssert.AreEqual(never, asNecessary,
			"A small archive should hold the same entries whether or not Zip64 is permitted.");
	}

	private static string[] ReadBackEntryNames(Zip64Option option)
	{
		using var stream = new MemoryStream();
		using (var zip = new ZipOutputStream(stream, leaveOpen: true) { EnableZip64 = option })
		{
			WriteEntries(zip, 3);
		}

		stream.Position = 0;
		using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
		return [.. archive.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal)];
	}

	private static void WriteEntries(ZipOutputStream zip, int count)
	{
		var payload = Encoding.UTF8.GetBytes("x");
		for (var i = 0; i < count; i++)
		{
			zip.PutNextEntry($"entry-{i}.txt");
			zip.Write(payload, 0, payload.Length);
		}
	}
}

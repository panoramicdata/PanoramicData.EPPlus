using OfficeOpenXml;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace PanoramicData.EPPlus.Test;

[TestClass]
public class PackageRoundTripTests
{
	private static readonly XNamespace _contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
	private static readonly XNamespace _relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

	private static string WorkbooksDirectory
		=> Path.Combine(Directory.GetParent(AppContext.BaseDirectory).Parent.Parent.Parent.FullName, "Workbooks");

	public static IEnumerable<object[]> Workbooks
		=> Directory.Exists(WorkbooksDirectory)
			? Directory.GetFiles(WorkbooksDirectory)
				.Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
				.OrderBy(f => f, StringComparer.Ordinal)
				.Select(f => new object[] { Path.GetFileName(f) })
			: [];

	[TestMethod]
	public void WorkbooksFolderHoldsTheRoundTripInputs()
		=> Assert.IsGreaterThanOrEqualTo(3, Workbooks.Count(), $"Expected at least 3 workbooks in {WorkbooksDirectory}");

	[TestMethod]
	[DynamicData(nameof(Workbooks))]
	public void Workbook_SaveAndReopen_KeepsSheetsCellsPartsAndContentTypes(string fileName)
	{
		var path = Path.Combine(WorkbooksDirectory, fileName);
		var original = File.ReadAllBytes(path);

		byte[] saved;
		List<string> originalCells;
		using (var package = new ExcelPackage(new FileInfo(path)))
		{
			originalCells = DumpCells(package);
			saved = package.GetAsByteArray();
		}

		using (var reopened = new ExcelPackage(new MemoryStream(saved)))
		{
			CollectionAssert.AreEqual(originalCells, DumpCells(reopened));
		}

		var originalParts = ReadParts(original);
		var savedParts = ReadParts(saved);
		var dropped = originalParts.Keys.Except(savedParts.Keys).ToList();
		var added = savedParts.Keys.Except(originalParts.Keys).ToList();

		Assert.IsTrue(dropped.All(name => IsDroppedBySave(name, original)), $"Parts lost on save: {string.Join(", ", dropped)}");
		Assert.IsTrue(added.All(name => name == "xl/sharedStrings.xml"), $"Unexpected parts added on save: {string.Join(", ", added)}");
		foreach (var name in originalParts.Keys.Intersect(savedParts.Keys))
		{
			Assert.AreEqual(originalParts[name], savedParts[name], $"Content type of {name}");
		}
	}

	[TestMethod]
	[DataRow(EncryptionVersion.Standard)]
	[DataRow(EncryptionVersion.Agile)]
	public void EncryptedPackage_RoundTripsThroughBytesAndStream(EncryptionVersion version)
	{
		var encrypted = CreateEncryptedPackage(version);

		Assert.AreEqual(0xD0, encrypted[0], "An encrypted package is an OLE compound document");
		using var package = new ExcelPackage(new MemoryStream(encrypted), "secret");
		var sheet = package.Workbook.Worksheets["Data"];
		Assert.AreEqual("hidden", sheet.Cells["A1"].Value);
		Assert.AreEqual(42.5, sheet.Cells["B2"].Value);
		Assert.IsTrue(package.Encryption.IsEncrypted);
	}

	[TestMethod]
	public void EncryptedPackage_RoundTripsThroughFile()
	{
		var path = Path.Combine(Path.GetTempPath(), $"epplus-encrypted-{Guid.NewGuid():N}.xlsx");
		try
		{
			File.WriteAllBytes(path, CreateEncryptedPackage(EncryptionVersion.Agile));
			using var package = new ExcelPackage(new FileInfo(path), "secret");
			Assert.AreEqual("hidden", package.Workbook.Worksheets["Data"].Cells["A1"].Value);
		}
		finally
		{
			File.Delete(path);
		}
	}

	[TestMethod]
	public void EncryptedPackage_WrongPasswordIsRejected()
		=> Assert.ThrowsExactly<UnauthorizedAccessException>(() => new ExcelPackage(new MemoryStream(CreateEncryptedPackage(EncryptionVersion.Standard)), "wrong"));

	[TestMethod]
	public void EncryptedPackage_WithoutPasswordIsNotAPackage()
	{
		var exception = Assert.ThrowsExactly<InvalidDataException>(() => new ExcelPackage(new MemoryStream(CreateEncryptedPackage(EncryptionVersion.Agile))));
		StringAssert.Contains(exception.Message, "please supply the password");
	}

	[TestMethod]
	public void NonZipStream_IsNotAPackage()
		=> Assert.ThrowsExactly<InvalidDataException>(() => new ExcelPackage(new MemoryStream([1, 2, 3, 4, 5, 6, 7, 8])));

	[TestMethod]
	public void SaveAs_LeavesTheCallersStreamsOpen()
	{
		var input = new MemoryStream(File.ReadAllBytes(Path.Combine(WorkbooksDirectory, "FormulaTest.xlsx")));
		var output = new MemoryStream();
		using (var package = new ExcelPackage(input))
		{
			package.SaveAs(output);
		}

		Assert.IsTrue(input.CanRead);
		Assert.IsTrue(output.CanWrite);
		output.Position = 0;
		using var reopened = new ExcelPackage(output);
		Assert.HasCount(4, reopened.Workbook.Worksheets);
	}

	[TestMethod]
	public void CompressionNone_StillDeflatesEntries()
	{
		using var package = new ExcelPackage();
		package.Compression = OfficeOpenXml.CompressionLevel.None;
		var sheet = package.Workbook.Worksheets.Add("Sheet");
		for (var row = 1; row <= 500; row++)
		{
			sheet.Cells[row, 1].Value = "row " + row;
		}

		using var archive = new ZipArchive(new MemoryStream(package.GetAsByteArray()), ZipArchiveMode.Read);
		var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
		Assert.IsLessThan(entry.Length, entry.CompressedLength);
	}

	private static byte[] CreateEncryptedPackage(EncryptionVersion version)
	{
		using var package = new ExcelPackage();
		var sheet = package.Workbook.Worksheets.Add("Data");
		sheet.Cells["A1"].Value = "hidden";
		sheet.Cells["B2"].Value = 42.5;
		package.Encryption.Version = version;
		package.Encryption.Password = "secret";
		return package.GetAsByteArray();
	}

	private static List<string> DumpCells(ExcelPackage package)
	{
		var cells = new List<string>();
		foreach (var sheet in package.Workbook.Worksheets)
		{
			cells.Add($"sheet {sheet.Name} {sheet.Dimension?.Address}");
			if (sheet.Dimension == null)
			{
				continue;
			}

			foreach (var cell in sheet.Cells[sheet.Dimension.Address])
			{
				cells.Add($"{cell.Address}={Convert.ToString(cell.Value, CultureInfo.InvariantCulture)}|{cell.Value?.GetType().Name}");
			}
		}

		return cells;
	}

	private static Dictionary<string, string> ReadParts(byte[] package)
	{
		using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
		var contentTypes = Load(archive, "[Content_Types].xml").Root;
		var defaults = contentTypes.Elements(_contentTypesNamespace + "Default")
			.ToDictionary(e => e.Attribute("Extension").Value, e => e.Attribute("ContentType").Value, StringComparer.OrdinalIgnoreCase);
		var overrides = contentTypes.Elements(_contentTypesNamespace + "Override")
			.ToDictionary(e => e.Attribute("PartName").Value.TrimStart('/'), e => e.Attribute("ContentType").Value, StringComparer.OrdinalIgnoreCase);

		return archive.Entries
			.Where(e => !e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
			.ToDictionary(
				e => e.FullName,
				e => overrides.TryGetValue(e.FullName, out var type) ? type : defaults.GetValueOrDefault(Path.GetExtension(e.FullName).TrimStart('.'), ""));
	}

	// EPPlus drops the calculation chain and printer settings on save, and a rels file left with nothing else to point at.
	private static bool IsDroppedBySave(string name, byte[] original)
	{
		if (name == "xl/calcChain.xml" || name.StartsWith("xl/printerSettings/", StringComparison.Ordinal))
		{
			return true;
		}

		if (!name.EndsWith(".rels", StringComparison.Ordinal))
		{
			return false;
		}

		using var archive = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read);
		return Load(archive, name).Root.Elements(_relationshipsNamespace + "Relationship")
			.All(r => r.Attribute("Target").Value.Contains("printerSettings", StringComparison.Ordinal));
	}

	private static XDocument Load(ZipArchive archive, string name)
	{
		using var stream = archive.GetEntry(name).Open();
		return XDocument.Load(stream);
	}
}

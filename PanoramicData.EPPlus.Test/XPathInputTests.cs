using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// OPS-161211: caller-supplied names and addresses must not change the XPath queries built from them.
/// </summary>
[TestClass]
public class XPathInputTests
{
	private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

	private static readonly byte[] OnePixelPng = Convert.FromBase64String(
		"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

	[TestMethod]
	public void CustomProperty_NameWithMixedQuotes_RoundTrips()
	{
		using var package = new ExcelPackage();
		var properties = package.Workbook.Properties;
		const string name = "O'Brien said \"hello\"";

		properties.SetCustomPropertyValue(name, "first");
		properties.SetCustomPropertyValue(name, "second");

		Assert.AreEqual("second", properties.GetCustomPropertyValue(name));
		Assert.AreEqual(1, CountCustomProperties(properties));
	}

	[TestMethod]
	public void CustomProperty_InjectedName_DoesNotOverwriteAnotherProperty()
	{
		using var package = new ExcelPackage();
		var properties = package.Workbook.Properties;
		const string injected = "z' or 'a'='a";

		properties.SetCustomPropertyValue("Victim", 1);
		Assert.IsNull(properties.GetCustomPropertyValue(injected));

		properties.SetCustomPropertyValue(injected, 2);

		Assert.AreEqual(1, properties.GetCustomPropertyValue("Victim"));
		Assert.AreEqual(2, properties.GetCustomPropertyValue(injected));
		Assert.AreEqual(2, CountCustomProperties(properties));
	}

	[TestMethod]
	public void ExtendedProperty_ValidName_RoundTrips()
	{
		using var package = new ExcelPackage();
		var properties = package.Workbook.Properties;

		properties.SetExtendedPropertyValue("Template", "Normal");

		Assert.AreEqual("Normal", properties.GetExtendedPropertyValue("Template"));
	}

	[TestMethod]
	public void ExtendedProperty_GetWithXPathSyntax_ReturnsNull()
	{
		using var package = new ExcelPackage();
		var properties = package.Workbook.Properties;
		properties.SetExtendedPropertyValue("Template", "Normal");

		Assert.IsNull(properties.GetExtendedPropertyValue("*"));
		Assert.IsNull(properties.GetExtendedPropertyValue("Template' or '1'='1"));
	}

	[TestMethod]
	public void ExtendedProperty_SetWithXPathSyntax_Throws()
	{
		using var package = new ExcelPackage();
		var properties = package.Workbook.Properties;

		Assert.ThrowsExactly<ArgumentException>(() => properties.SetExtendedPropertyValue("Template[1]", "x"));
		Assert.ThrowsExactly<ArgumentException>(() => properties.SetExtendedPropertyValue("a'b", "x"));
	}

	[TestMethod]
	public void CopyWorksheet_PictureNameWithMixedQuotes_CopiesThePicture()
	{
		const string pictureName = "O'Brien's \"logo\"";
		byte[] saved;
		using (var package = new ExcelPackage())
		{
			var source = package.Workbook.Worksheets.Add("Source");
			// Removing the first picture leaves the copied one on rId2, so the copy only resolves if its r:embed is rewritten to the new rId1.
			source.Drawings.AddPicture("Removed", EPPlusTest.Properties.Resources.Test1, 10, 10);
			source.Drawings.AddPicture(pictureName, OnePixelPng, 10, 10);
			source.Drawings.Remove("Removed");

			package.Workbook.Worksheets.Add("Copy", source);

			saved = package.GetAsByteArray();
		}

		using var reopened = new ExcelPackage(new MemoryStream(saved));
		var copied = reopened.Workbook.Worksheets["Copy"].Drawings[pictureName] as ExcelPicture;
		Assert.IsNotNull(copied);
		CollectionAssert.AreEqual(OnePixelPng, copied.ImageBytes);
	}

	[TestMethod]
	public void ConditionalFormatting_AddressThatWouldEndTheXPathLiteral_IsRejected()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("Data");
		worksheet.ConditionalFormatting.AddEqual(new ExcelAddress("C3:C9"));

		var injected = new ExcelAddress("' or 1=1 or '!A1");

		Assert.ThrowsExactly<ArgumentException>(() => worksheet.ConditionalFormatting.AddEqual(injected));

		var namespaces = new XmlNamespaceManager(worksheet.WorksheetXml.NameTable);
		namespaces.AddNamespace("d", MainNamespace);
		var sqrefs = worksheet.WorksheetXml.SelectNodes("//d:conditionalFormatting/@sqref", namespaces);
		Assert.AreEqual(1, sqrefs.Count);
		Assert.AreEqual("C3:C9", sqrefs[0].Value);
	}

	[TestMethod]
	public void CopyWorksheet_ChartOnALaterRelationshipId_IsRelinked()
	{
		byte[] saved;
		using (var package = new ExcelPackage())
		{
			var source = package.Workbook.Worksheets.Add("Source");
			source.Cells["A1"].Value = 1;
			source.Cells["A2"].Value = 2;
			source.Drawings.AddChart("Removed", eChartType.ColumnClustered).Series.Add("A1:A2", "A1:A2");
			source.Drawings.AddChart("Kept", eChartType.ColumnClustered).Series.Add("A1:A2", "A1:A2");
			source.Drawings.Remove("Removed");

			package.Workbook.Worksheets.Add("Copy", source);

			saved = package.GetAsByteArray();
		}

		using var reopened = new ExcelPackage(new MemoryStream(saved));
		var copied = reopened.Workbook.Worksheets["Copy"].Drawings["Kept"] as ExcelChart;
		Assert.IsNotNull(copied);
		Assert.AreEqual(1, copied.Series.Count);
	}

	private static int CountCustomProperties(OfficeProperties properties)
		=> properties.CustomPropertiesXml.DocumentElement.ChildNodes.Count;
}

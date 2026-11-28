using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace PanoramicData.EPPlus.Test;

/// <summary>
/// OPS-161211: caller-supplied names and addresses must not change XPath queries, and lookups rewritten to use fixed XPath keep their behaviour.
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

	[TestMethod]
	public void StyleFlags_SurviveSaveAndReload()
	{
		byte[] saved;
		using (var package = new ExcelPackage())
		{
			var worksheet = package.Workbook.Worksheets.Add("Styles");
			var first = worksheet.Cells["A1"].Style;
			first.Font.Bold = true;
			first.Font.Strike = true;
			first.Border.DiagonalUp = true;
			var second = worksheet.Cells["A2"].Style;
			second.Font.Italic = true;
			second.Border.DiagonalDown = true;
			worksheet.Cells["A3"].Value = "plain";
			saved = package.GetAsByteArray();
		}

		using var reopened = new ExcelPackage(new MemoryStream(saved));
		var sheet = reopened.Workbook.Worksheets["Styles"];
		AssertFlags(sheet.Cells["A1"].Style, bold: true, italic: false, strike: true, diagonalUp: true, diagonalDown: false);
		AssertFlags(sheet.Cells["A2"].Style, bold: false, italic: true, strike: false, diagonalUp: false, diagonalDown: true);
		AssertFlags(sheet.Cells["A3"].Style, bold: false, italic: false, strike: false, diagonalUp: false, diagonalDown: false);

		static void AssertFlags(OfficeOpenXml.Style.ExcelStyle style, bool bold, bool italic, bool strike, bool diagonalUp, bool diagonalDown)
		{
			Assert.AreEqual(bold, style.Font.Bold, "bold");
			Assert.AreEqual(italic, style.Font.Italic, "italic");
			Assert.AreEqual(strike, style.Font.Strike, "strike");
			Assert.AreEqual(diagonalUp, style.Border.DiagonalUp, "diagonalUp");
			Assert.AreEqual(diagonalDown, style.Border.DiagonalDown, "diagonalDown");
		}
	}

	[TestMethod]
	public void ColorScaleTypes_SurviveSaveAndReload()
	{
		byte[] saved;
		using (var package = new ExcelPackage())
		{
			var worksheet = package.Workbook.Worksheets.Add("Scales");
			var three = worksheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:A5"));
			three.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Num;
			three.MiddleValue.Value = 50;
			three.HighValue.Type = eExcelConditionalFormattingValueObjectType.Num;
			three.HighValue.Value = 90;
			var two = worksheet.ConditionalFormatting.AddTwoColorScale(new ExcelAddress("B1:B5"));
			two.HighValue.Type = eExcelConditionalFormattingValueObjectType.Num;
			two.HighValue.Value = 80;
			saved = package.GetAsByteArray();
		}

		using var reopened = new ExcelPackage(new MemoryStream(saved));
		var rules = reopened.Workbook.Worksheets["Scales"].ConditionalFormatting;
		Assert.AreEqual(1, rules.Count(rule => rule.Type == eExcelConditionalFormattingRuleType.ThreeColorScale));
		Assert.AreEqual(1, rules.Count(rule => rule.Type == eExcelConditionalFormattingRuleType.TwoColorScale));
	}

	[TestMethod]
	public void GetCfvoNodeByPosition_PicksTheCfvoByOrder()
	{
		var document = new XmlDocument { XmlResolver = null };
		document.LoadXml($"<colorScale xmlns=\"{MainNamespace}\"><cfvo type=\"min\"/><cfvo type=\"percentile\"/><cfvo type=\"max\"/></colorScale>");
		var namespaces = new XmlNamespaceManager(document.NameTable);
		namespaces.AddNamespace("d", MainNamespace);
		var top = document.DocumentElement;

		Assert.AreEqual("min", CfvoType(eExcelConditionalFormattingValueObjectPosition.Low, eExcelConditionalFormattingRuleType.ThreeColorScale));
		Assert.AreEqual("percentile", CfvoType(eExcelConditionalFormattingValueObjectPosition.Middle, eExcelConditionalFormattingRuleType.ThreeColorScale));
		Assert.AreEqual("max", CfvoType(eExcelConditionalFormattingValueObjectPosition.High, eExcelConditionalFormattingRuleType.ThreeColorScale));
		Assert.AreEqual("percentile", CfvoType(eExcelConditionalFormattingValueObjectPosition.High, eExcelConditionalFormattingRuleType.TwoColorScale));

		top.RemoveChild(top.LastChild);
		Assert.ThrowsExactly<InvalidOperationException>(() => CfvoType(eExcelConditionalFormattingValueObjectPosition.High, eExcelConditionalFormattingRuleType.ThreeColorScale));

		string CfvoType(eExcelConditionalFormattingValueObjectPosition position, eExcelConditionalFormattingRuleType ruleType)
			=> ExcelConditionalFormattingValueObjectType.GetCfvoNodeByPosition(position, ruleType, top, namespaces).Attributes["type"].Value;
	}

	[TestMethod]
	public void ChartDataLabelsAndDataTable_AreInsertedAfterTheirSchemaPredecessors()
	{
		using var package = new ExcelPackage();
		var worksheet = package.Workbook.Worksheets.Add("Chart");
		worksheet.Cells["A1"].Value = 1;
		worksheet.Cells["A2"].Value = 2;
		var chart = (ExcelBarChart)worksheet.Drawings.AddChart("Bars", eChartType.ColumnClustered);
		chart.Series.Add("A1:A2", "A1:A2");

		chart.DataLabel.ShowValue = true;
		chart.PlotArea.CreateDataTable();

		var namespaces = new XmlNamespaceManager(chart.ChartXml.NameTable);
		namespaces.AddNamespace("c", ExcelPackage.schemaChart);
		var dataLabels = chart.ChartXml.SelectSingleNode("//c:barChart/c:dLbls", namespaces);
		var dataTable = chart.ChartXml.SelectSingleNode("//c:plotArea/c:dTable", namespaces);
		Assert.AreEqual("ser", dataLabels?.PreviousSibling?.LocalName);
		Assert.AreEqual("valAx", dataTable?.PreviousSibling?.LocalName);
	}

	private static int CountCustomProperties(OfficeProperties properties)
		=> properties.CustomPropertiesXml.DocumentElement.ChildNodes.Count;
}

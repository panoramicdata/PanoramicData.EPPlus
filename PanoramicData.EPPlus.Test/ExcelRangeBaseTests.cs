using OfficeOpenXml;

namespace PanoramicData.EPPlus.Test;

[TestClass]
public class ExcelRangeBaseTests : TestBase
{
	[TestMethod]
	public void CopyCopiesCommentsFromSingleCellRanges()
	{
		InitBase();
		var pck = new ExcelPackage();
		var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
		var sourceExcelRange = ws1.Cells[3, 3];
		Assert.IsNull(sourceExcelRange.Comment);
		sourceExcelRange.AddComment("Testing comment 1", "test1");
		Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
		Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
		var destinationExcelRange = ws1.Cells[5, 5];
		Assert.IsNull(destinationExcelRange.Comment);
		sourceExcelRange.Copy(destinationExcelRange);
		// Assert the original comment is intact.
		Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
		Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
		// Assert the comment was copied.
		Assert.AreEqual("test1", destinationExcelRange.Comment.Author);
		Assert.AreEqual("Testing comment 1", destinationExcelRange.Comment.Text);
	}

	[TestMethod]
	public void CopyCopiesCommentsFromMultiCellRanges()
	{
		InitBase();
		var pck = new ExcelPackage();
		var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
		var sourceExcelRangeC3 = ws1.Cells[3, 3];
		var sourceExcelRangeD3 = ws1.Cells[3, 4];
		var sourceExcelRangeE3 = ws1.Cells[3, 5];
		Assert.IsNull(sourceExcelRangeC3.Comment);
		Assert.IsNull(sourceExcelRangeD3.Comment);
		Assert.IsNull(sourceExcelRangeE3.Comment);
		sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
		sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
		sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
		AssertNumberedComments(sourceExcelRangeC3, sourceExcelRangeD3, sourceExcelRangeE3);
		// Copy the full row to capture each cell at once.
		Assert.IsNull(ws1.Cells[5, 3].Comment);
		Assert.IsNull(ws1.Cells[5, 4].Comment);
		Assert.IsNull(ws1.Cells[5, 5].Comment);
		ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);
		// Assert the original comments are intact.
		AssertNumberedComments(sourceExcelRangeC3, sourceExcelRangeD3, sourceExcelRangeE3);
		// Assert the comments were copied.
		var destinationExcelRangeC5 = ws1.Cells[5, 3];
		var destinationExcelRangeD5 = ws1.Cells[5, 4];
		var destinationExcelRangeE5 = ws1.Cells[5, 5];
		AssertNumberedComments(destinationExcelRangeC5, destinationExcelRangeD5, destinationExcelRangeE5);
	}

	private static void AssertNumberedComments(params ExcelRange[] cells)
	{
		for (var i = 0; i < cells.Length; i++)
		{
			Assert.AreEqual("test1", cells[i].Comment.Author);
			Assert.AreEqual($"Testing comment {i + 1}", cells[i].Comment.Text);
		}
	}

	[TestMethod]
	public void SettingAddressHandlesMultiAddresses()
	{
		using ExcelPackage package = new();
		var worksheet = package.Workbook.Worksheets.Add("Sheet1");
		var name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
		name.Address = "Sheet1!C3";
		name.Address = "Sheet1!D3";
		Assert.IsNull(name.Addresses);
		name.Address = "C3:D3,E3:F3";
		Assert.IsNotNull(name.Addresses);
		name.Address = "Sheet1!C3";
		Assert.IsNull(name.Addresses);
	}
}

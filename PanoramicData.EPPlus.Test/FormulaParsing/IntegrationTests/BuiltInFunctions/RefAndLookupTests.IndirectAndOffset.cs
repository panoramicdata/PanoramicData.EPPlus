using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class RefAndLookupTests
{
	[TestMethod]
	public void IndirectShouldReturnARange()
	{
		using var package = new ExcelPackage(new MemoryStream());
		var s1 = package.Workbook.Worksheets.Add("Test");
		s1.Cells["A1:A2"].Value = 2;
		s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
		s1.Calculate();
		Assert.AreEqual(4d, s1.Cells["A3"].Value);

		s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
		s1.Calculate();
		Assert.AreEqual(4d, s1.Cells["A4"].Value);
	}

	[TestMethod]
	public void OffsetShouldReturnASingleValue()
		=> Assert.AreEqual(1d, CalculateOnNewSheet("Test", "A5", "OFFSET(A1, 2, 1)", ("B3", 1d)));

	[TestMethod]
	public void OffsetShouldReturnARange()
		=> Assert.AreEqual(3d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:A3, 0, 1))", OnesInColumnB));

	[TestMethod]
	public void OffsetShouldReturnARange2()
		=> Assert.AreEqual(3d, CalculateOnNewSheet("Test", "A5", "COUNTA(OFFSET(Test!B1, 0, 0, Test!B2, Test!B3))",
			("B1", 10d), ("B2", 10d), ("B3", 10d)));

	[TestMethod]
	public void OffsetDirectReferenceToMultiRangeShouldSetValueError()
	{
		var result = CalculateOnNewSheet("Test", "A5", "OFFSET(A1:A3, 0, 1)", OnesInColumnB);
		Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
	}

	[TestMethod]
	public void OffsetShouldReturnARangeAccordingToWidth()
		=> Assert.AreEqual(2d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:A3, 0, 1, 2))", OnesInColumnB));

	[TestMethod]
	public void OffsetShouldReturnARangeAccordingToHeight()
		=> Assert.AreEqual(6d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:A3, 0, 1, 2, 2))",
			[.. OnesInColumnB, ("C1", 2d), ("C2", 2d), ("C3", 2d)]));

	[TestMethod]
	public void OffsetShouldCoverMultipleColumns()
		=> Assert.AreEqual(9d, CalculateOnNewSheet("Test", "A5", "SUM(OFFSET(A1:B3, 0, 2))",
			("C1", 1d), ("C2", 1d), ("C3", 1d), ("D1", 2d), ("D2", 2d), ("D3", 2d)));

	[TestMethod, Ignore]
	public void VLookupShouldHandleNames()
	{
		using var package = new ExcelPackage(new FileInfo(@"c:\temp\Book3.xlsx"));
		var s1 = package.Workbook.Worksheets.First();
		_ = s1.Cells["X10"].Formula;
	}

	[TestMethod]
	public void LookupShouldReturnFromResultVector()
	{
		var lookupAddress = "A1:A5";
		var resultAddress = "B1:B5";
		var result = CalculateOnNewSheet("test", "E5", "LOOKUP(C1, " + lookupAddress + ", " + resultAddress + ")",
			("A1", 4.14), ("A2", 4.19), ("A3", 5.17), ("A4", 5.77), ("A5", 6.39),
			("B1", "red"), ("B2", "orange"), ("B3", "yellow"), ("B4", "green"), ("B5", "blue"),
			("C1", 4.14));
		Assert.AreEqual("red", result);
	}
}

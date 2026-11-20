using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

public partial class RefAndLookupTests
{
	[TestMethod]
	public void LookupArgumentsShouldSetSearchedValue()
		=> Assert.AreEqual(1, CreateLookupArguments(1, "A:B", 2).SearchedValue);

	[TestMethod]
	public void LookupArgumentsShouldSetRangeAddress()
		=> Assert.AreEqual("A:B", CreateLookupArguments(1, "A:B", 2).RangeAddress);

	[TestMethod]
	public void LookupArgumentsShouldSetColIndex()
		=> Assert.AreEqual(2, CreateLookupArguments(1, "A:B", 2).LookupIndex);

	[TestMethod]
	public void LookupArgumentsShouldSetRangeLookupToTrueAsDefaultValue()
		=> Assert.IsTrue(CreateLookupArguments(1, "A:B", 2).RangeLookup);

	[TestMethod]
	public void LookupArgumentsShouldSetRangeLookupToTrueWhenTrueIsSupplied()
		=> Assert.IsTrue(CreateLookupArguments(1, "A:B", 2, true).RangeLookup);

	[TestMethod]
	public void VLookupShouldReturnResultFromMatchingRow()
	{
		var result = ExecuteOnFakeCells(new VLookup(), [2, "A1:B2", 2], true,
			(1, 1, 1), (1, 2, 1), (2, 1, 2), (2, 2, 5));
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
	{
		var result = ExecuteOnFakeCells(new VLookup(), [4, "A1:B2", 2, true], false,
			(1, 1, 3), (1, 2, 1), (2, 1, 5), (2, 2, 4));
		Assert.AreEqual(1, result.Result);
	}

	[TestMethod]
	public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
	{
		var result = ExecuteOnFakeCells(new VLookup(), ["B", "A1:B2", 2, true], false,
			(1, 1, "A"), (1, 2, 1), (2, 1, "C"), (2, 2, 4));
		Assert.AreEqual(1, result.Result);
	}
}

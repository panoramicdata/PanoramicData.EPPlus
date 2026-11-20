using OfficeOpenXml.FormulaParsing;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;
using Index = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Index;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

[TestClass]
public class IndexTests : WorksheetTestBase
{
	[TestMethod]
	public void Index_Should_Return_Value_By_Index()
	{
		var func = new Index();
		var result = func.Execute(
			FunctionsHelper.CreateArgs(
				FunctionsHelper.CreateArgs(1, 2, 5),
				3
				), Context);
		Assert.AreEqual(5, result.Result);
	}

	[TestMethod]
	public void Index_Should_Handle_SingleRange()
	{
		Worksheet.Cells["A1"].Value = 1d;
		Worksheet.Cells["A2"].Value = 3d;
		Worksheet.Cells["A3"].Value = 5d;

		Worksheet.Cells["A4"].Formula = "INDEX(A1:A3;3)";

		Worksheet.Calculate();

		Assert.AreEqual(5d, Worksheet.Cells["A4"].Value);
	}
}

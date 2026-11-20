using OfficeOpenXml.FormulaParsing;
using PanoramicData.EPPlus.Test.FormulaParsing.TestHelpers;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.RefAndLookup;

[TestClass]
public class ChooseTests : WorksheetTestBase
{
	[TestMethod]
	public void ChooseSingleValue()
	{
		fillChooseOptions();
		Worksheet.Cells["B1"].Formula = "CHOOSE(4, A1, A2, A3, A4, A5)";
		Worksheet.Calculate();

		Assert.AreEqual("5", Worksheet.Cells["B1"].Value);
	}

	[TestMethod]
	public void ChooseSingleFormula()
	{
		fillChooseOptions();
		Worksheet.Cells["B1"].Formula = "CHOOSE(6, A1, A2, A3, A4, A5, A6)";
		Worksheet.Calculate();

		Assert.AreEqual("12", Worksheet.Cells["B1"].Value);
	}

	[TestMethod]
	public void ChooseMultipleValues()
	{
		fillChooseOptions();
		Worksheet.Cells["B1"].Formula = "SUM(CHOOSE({1,3,4}, A1, A2, A3, A4, A5))";
		Worksheet.Calculate();

		Assert.AreEqual(9D, Worksheet.Cells["B1"].Value);
	}

	[TestMethod]
	public void ChooseValueAndFormula()
	{
		fillChooseOptions();
		Worksheet.Cells["B1"].Formula = "SUM(CHOOSE({2,6}, A1, A2, A3, A4, A5, A6))";
		Worksheet.Calculate();

		Assert.AreEqual(14D, Worksheet.Cells["B1"].Value);
	}

	private void fillChooseOptions()
	{
		Worksheet.Cells["A1"].Value = 1d;
		Worksheet.Cells["A2"].Value = 2d;
		Worksheet.Cells["A3"].Value = 3d;
		Worksheet.Cells["A4"].Value = 5d;
		Worksheet.Cells["A5"].Value = 7d;
		Worksheet.Cells["A6"].Formula = "A4 + A5";
	}
}

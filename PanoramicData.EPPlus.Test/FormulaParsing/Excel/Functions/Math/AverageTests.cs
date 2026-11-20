using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class AverageTests
{
	[TestMethod]
	public void AverageLiterals()
	{
		// In the case of literals, Average DOES parse and include numeric strings, date strings, bools, etc.
		Average average = new();
		var result = AverageTestData.ExecuteOnLiterals(average);
		Assert.AreEqual(AverageTestData.ExpectedLiteralAverage, result.Result);
	}

	[TestMethod]
	public void AverageCellReferences()
	{
		// In the case of cell references, Average DOES NOT parse and include numeric strings, date strings, bools, unparsable strings, etc.
		Average average = new();
		var result = AverageTestData.ExecuteOnCellReferences(average);
		Assert.AreEqual((2000 + new DateTime(2013, 1, 5).ToOADate()) / 2, result.Result);
	}

	[TestMethod]
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S3257:Declarations and initializations should be as concise as possible", Justification = "The explicit array type is what target-types the new() elements; without it the code does not compile.")]
	public void AverageArray()
	{
		// In the case of arrays, Average DOES NOT parse and include numeric strings, date strings, bools, unparsable strings, etc.
		Average average = new();
		var date1 = new DateTime(2013, 1, 5);
		var date2 = new DateTime(2013, 1, 15);
		double value = 2000;
		var result = average.Execute(
		[
			new(new FunctionArgument[]
			{
				new(1000.ToString("n")),
				new(value),
				new(6000.ToString("n")),
				new(true),
				new(date1),
				new(date2.ToString("d")),
				new("test")
			})
		], ParsingContext.Create());
		Assert.AreEqual((2000 + date1.ToOADate()) / 2, result.Result);
	}

	[TestMethod]
	public void AverageUnparsableLiteral()
	{
		// In the case of literals, any unparsable string literal results in a #VALUE.
		Average average = new();
		Assert.ThrowsExactly<ExcelErrorValueException>(() => average.Execute(
		[
			new(1000),
			new("Test")
		], ParsingContext.Create()));
	}
}

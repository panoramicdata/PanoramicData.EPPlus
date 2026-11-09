using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions.Math;

[TestClass]
public class AverageATests
{
	[TestMethod]
	public void AverageALiterals()
	{
		// For literals, AverageA always parses and include numeric strings, date strings, bools, etc.
		// The only exception is unparsable string literals, which cause a #VALUE.
		AverageA average = new();
		var result = AverageTestData.ExecuteOnLiterals(average);
		Assert.AreEqual(AverageTestData.ExpectedLiteralAverage, result.Result);
	}

	[TestMethod]
	public void AverageACellReferences()
	{
		// For cell references, AverageA divides by all cells, but only adds actual numbers, dates, and booleans.
		double[] values =
		[
			0,
			2000,
			0,
			1,
			new DateTime(2013, 1, 5).ToOADate(),
			0
		];
		AverageA average = new();
		var result = AverageTestData.ExecuteOnCellReferences(average);
		Assert.AreEqual(values.Average(), result.Result);
	}

	[TestMethod]
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S3257:Declarations and initializations should be as concise as possible", Justification = "The explicit array type is what target-types the new() elements; without it the code does not compile.")]
	public void AverageAArray()
	{
		// For arrays, AverageA completely ignores booleans.  It divides by strings and numbers, but only
		// numbers are added to the total.  Real dates cannot be specified and string dates are not parsed.
		AverageA average = new();
		var date = new DateTime(2013, 1, 15);
		double[] values =
		[
			0,
			2000,
			0,
			0,
			0
		];
		var result = average.Execute(
		[
			new(new FunctionArgument[]
			{
				new(1000.ToString("n")),
				new(2000),
				new(6000.ToString("n")),
				new(true),
				new(date.ToString("d")),
				new("test")
			})
		], ParsingContext.Create());
		Assert.AreEqual(values.Average(), result.Result);
	}

	[TestMethod]
	public void AverageAUnparsableLiteral()
	{
		// In the case of literals, any unparsable string literal results in a #VALUE.
		AverageA average = new();
		Assert.ThrowsExactly<ExcelErrorValueException>(() => average.Execute(
		[
			new(1000),
			new("Test")
		], ParsingContext.Create()));
	}
}

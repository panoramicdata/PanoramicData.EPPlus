using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel;

public partial class OperatorsTests
{
	[TestMethod]
	public void OperatorConcatShouldConcatTwoStrings()
	{
		var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String));
		Assert.AreEqual("ab", result.Result);
	}

	[TestMethod]
	public void OperatorConcatShouldConcatANumberAndAString()
	{
		var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String));
		Assert.AreEqual("12b", result.Result);
	}

	[TestMethod]
	public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
	{
		var result = Operator.Eq.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer));
		Assert.IsTrue((bool)result.Result);
	}

	[TestMethod]
	public void OperatorEqShouldReturnFalsefSuppliedValuesDiffer()
	{
		var result = Operator.Eq.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
		Assert.IsFalse((bool)result.Result);
	}

	[TestMethod]
	public void OperatorNotEqualToShouldReturnTruefSuppliedValuesDiffer()
	{
		var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
		Assert.IsTrue((bool)result.Result);
	}

	[TestMethod]
	public void OperatorNotEqualToShouldReturnFalsefSuppliedValuesAreEqual()
	{
		var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(11, DataType.Integer));
		Assert.IsFalse((bool)result.Result);
	}

	[TestMethod]
	public void OperatorGreaterThanToShouldReturnTrueIfLeftIsSetAndRightIsNull()
	{
		var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(null, DataType.Empty));
		Assert.IsTrue((bool)result.Result);
	}

	[TestMethod]
	public void OperatorGreaterThanToShouldReturnTrueIfLeftIs11AndRightIs10()
	{
		var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(10, DataType.Integer));
		Assert.IsTrue((bool)result.Result);
	}

	[TestMethod]
	public void OperatorExpShouldReturnCorrectResult()
	{
		var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer));
		Assert.AreEqual(8d, result.Result);
	}

	[TestMethod]
	public void OperatorsActingOnNumericStrings()
	{
		var number1 = 42.0;
		var number2 = -143.75;
		CompileResult result1 = new(number1.ToString("n"), DataType.String);
		CompileResult result2 = new(number2.ToString("n"), DataType.String);
		var operatorResult = Operator.Concat.Apply(result1, result2);
		Assert.AreEqual($"{number1:n}{number2:n}", operatorResult.Result);
		operatorResult = Operator.Divide.Apply(result1, result2);
		Assert.AreEqual(number1 / number2, operatorResult.Result);
		operatorResult = Operator.Exp.Apply(result1, result2);
		Assert.AreEqual(Math.Pow(number1, number2), operatorResult.Result);
		operatorResult = Operator.Minus.Apply(result1, result2);
		Assert.AreEqual(number1 - number2, operatorResult.Result);
		operatorResult = Operator.Multiply.Apply(result1, result2);
		Assert.AreEqual(number1 * number2, operatorResult.Result);
		operatorResult = Operator.Percent.Apply(result1, result2);
		Assert.AreEqual(number1 * number2, operatorResult.Result);
		operatorResult = Operator.Plus.Apply(result1, result2);
		Assert.AreEqual(number1 + number2, operatorResult.Result);
		// Comparison operators always compare string-wise and don't parse out the actual numbers.
		operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
		Assert.IsTrue((bool)operatorResult.Result);
		operatorResult = Operator.Eq.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
		Assert.IsFalse((bool)operatorResult.Result);
		AssertFirstIsGreater(result1, result2);
	}

	[TestMethod]
	public void OperatorsActingOnDateStrings()
	{
		const string dateFormat = "M-dd-yyyy";
		DateTime date1 = new(2015, 2, 20);
		DateTime date2 = new(2015, 12, 1);
		var numericDate1 = date1.ToOADate();
		var numericDate2 = date2.ToOADate();
		CompileResult result1 = new(date1.ToString(dateFormat), DataType.String); // 2/20/2015
		CompileResult result2 = new(date2.ToString(dateFormat), DataType.String); // 12/1/2015
		var operatorResult = Operator.Concat.Apply(result1, result2);
		Assert.AreEqual($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", operatorResult.Result);
		operatorResult = Operator.Divide.Apply(result1, result2);
		Assert.AreEqual(numericDate1 / numericDate2, operatorResult.Result);
		operatorResult = Operator.Exp.Apply(result1, result2);
		Assert.AreEqual(Math.Pow(numericDate1, numericDate2), operatorResult.Result);
		operatorResult = Operator.Minus.Apply(result1, result2);
		Assert.AreEqual(numericDate1 - numericDate2, operatorResult.Result);
		operatorResult = Operator.Multiply.Apply(result1, result2);
		Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
		operatorResult = Operator.Percent.Apply(result1, result2);
		Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
		operatorResult = Operator.Plus.Apply(result1, result2);
		Assert.AreEqual(numericDate1 + numericDate2, operatorResult.Result);
		// Comparison operators always compare string-wise and don't parse out the actual numbers.
		operatorResult = Operator.Eq.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
		Assert.IsFalse((bool)operatorResult.Result);
		operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
		Assert.IsTrue((bool)operatorResult.Result);
		AssertFirstIsGreater(result1, result2);
	}

	private static void AssertFirstIsGreater(CompileResult result1, CompileResult result2)
	{
		var operatorResult = Operator.GreaterThan.Apply(result1, result2);
		Assert.IsTrue((bool)operatorResult.Result);
		operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
		Assert.IsTrue((bool)operatorResult.Result);
		operatorResult = Operator.LessThan.Apply(result1, result2);
		Assert.IsFalse((bool)operatorResult.Result);
		operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
		Assert.IsFalse((bool)operatorResult.Result);
	}
}

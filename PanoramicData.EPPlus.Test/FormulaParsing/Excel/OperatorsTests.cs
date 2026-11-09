using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel;

[TestClass]
public partial class OperatorsTests
{
	[TestMethod]
	public void OperatorPlusShouldThrowExceptionIfNonNumericOperand()
	{
		var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
		Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
	}

	[TestMethod]
	public void OperatorPlusShouldAddNumericStringAndNumber()
	{
		var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("2", DataType.String));
		Assert.AreEqual(3d, result.Result);
	}

	[TestMethod]
	public void OperatorMinusShouldThrowExceptionIfNonNumericOperand()
	{
		var result = Operator.Minus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
		Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
	}

	[TestMethod]
	public void OperatorMinusShouldSubtractNumericStringAndNumber()
	{
		var result = Operator.Minus.Apply(new CompileResult(5, DataType.Integer), new CompileResult("2", DataType.String));
		Assert.AreEqual(3d, result.Result);
	}

	[TestMethod]
	public void OperatorDivideShouldReturnDivideByZeroIfRightOperandIsZero()
	{
		var result = Operator.Divide.Apply(new CompileResult(1d, DataType.Decimal), new CompileResult(0d, DataType.Decimal));
		Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result.Result);
	}

	[TestMethod]
	public void OperatorDivideShouldDivideCorrectly()
	{
		var result = Operator.Divide.Apply(new CompileResult(9d, DataType.Decimal), new CompileResult(3d, DataType.Decimal));
		Assert.AreEqual(3d, result.Result);
	}

	[TestMethod]
	public void OperatorDivideShouldReturnValueErrorIfNonNumericOperand()
	{
		var result = Operator.Divide.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
		Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
	}

	[TestMethod]
	public void OperatorDivideShouldDivideNumericStringAndNumber()
	{
		var result = Operator.Divide.Apply(new CompileResult(9, DataType.Integer), new CompileResult("3", DataType.String));
		Assert.AreEqual(3d, result.Result);
	}

	[TestMethod]
	public void OperatorMultiplyShouldThrowExceptionIfNonNumericOperand() => Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));

	[TestMethod]
	public void OperatoMultiplyShouldMultiplyNumericStringAndNumber()
	{
		var result = Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("3", DataType.String));
		Assert.AreEqual(3d, result.Result);
	}
}

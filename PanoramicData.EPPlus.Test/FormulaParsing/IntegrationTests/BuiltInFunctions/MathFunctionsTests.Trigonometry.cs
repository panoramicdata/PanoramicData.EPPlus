
namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void CosShouldReturnAResult()
	{
		var result = _parser.Parse("Cos(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void CoshShouldReturnAResult()
	{
		var result = _parser.Parse("Cosh(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void SinShouldReturnAResult()
	{
		var result = _parser.Parse("Sin(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void SinhShouldReturnAResult()
	{
		var result = _parser.Parse("Sinh(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void TanShouldReturnAResult()
	{
		var result = _parser.Parse("Tan(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void AtanShouldReturnAResult()
	{
		var result = _parser.Parse("Atan(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void Atan2ShouldReturnAResult()
	{
		var result = _parser.Parse("Atan2(2,1)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void TanhShouldReturnAResult()
	{
		var result = _parser.Parse("Tanh(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void LogShouldReturnAResult()
	{
		var result = _parser.Parse("Log(2, 2)");
		Assert.AreEqual(1d, result);
	}

	[TestMethod]
	public void Log10ShouldReturnAResult()
	{
		var result = _parser.Parse("Log10(2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void LnShouldReturnAResult()
	{
		var result = _parser.Parse("Ln(2)");
		Assert.IsInstanceOfType<double>(result);
	}
}

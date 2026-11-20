
namespace PanoramicData.EPPlus.Test.FormulaParsing.IntegrationTests.BuiltInFunctions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void RandShouldReturnAResult()
	{
		var result = _parser.Parse("Rand()");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void RandBetweenShouldReturnAResult()
	{
		var result = _parser.Parse("RandBetween(1,2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void CountShouldReturnAResult()
	{
		var result = _parser.Parse("Count(1,2,2,\"4\")");
		Assert.AreEqual(4d, result);
	}

	[TestMethod]
	public void CountAShouldReturnAResult()
	{
		var result = _parser.Parse("CountA(1,2,2,\"\", \"a\")");
		Assert.AreEqual(4d, result);
	}

	[TestMethod]
	public void CountIfShouldReturnAResult()
	{
		var result = _parser.Parse("CountIf({1;2;2;\"\"}, \"2\")");
		Assert.AreEqual(2d, result);
	}

	[TestMethod]
	public void VarShouldReturnAResult()
	{
		var result = _parser.Parse("Var(1,2,3)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void VarPShouldReturnAResult()
	{
		var result = _parser.Parse("VarP(1,2,3)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void ModShouldReturnAResult()
	{
		var result = _parser.Parse("Mod(5,2)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void SubtotalShouldReturnAResult()
	{
		var result = _parser.Parse("Subtotal(1, 10, 20)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void TruncShouldReturnAResult()
	{
		var result = _parser.Parse("Trunc(1.2345)");
		Assert.IsInstanceOfType<double>(result);
	}

	[TestMethod]
	public void ProductShouldReturnAResult()
	{
		var result = _parser.Parse("Product(1,2,3)");
		Assert.IsInstanceOfType<double>(result);
	}
}

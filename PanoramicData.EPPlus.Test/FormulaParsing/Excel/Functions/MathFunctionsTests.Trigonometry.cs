using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace PanoramicData.EPPlus.Test.FormulaParsing.Excel.Functions;

public partial class MathFunctionsTests
{
	[TestMethod]
	public void ModShouldReturnCorrectResult() => Assert.AreEqual(1d, ExecuteFunction(new Mod(), 5, 2));

	[TestMethod]
	public void CosShouldReturnCorrectResult() => Assert.AreEqual(-0.416146837d, ExecuteFunctionRounded(new Cos(), 9, 2));

	[TestMethod]
	public void CosHShouldReturnCorrectResult() => Assert.AreEqual(3.762195691, ExecuteFunctionRounded(new Cosh(), 9, 2));

	[TestMethod]
	public void AcosShouldReturnCorrectResult() => Assert.AreEqual(1.4706, ExecuteFunctionRounded(new Acos(), 4, 0.1));

	[TestMethod]
	public void ACosHShouldReturnCorrectResult() => Assert.AreEqual(1.317, ExecuteFunctionRounded(new Acosh(), 3, 2));

	[TestMethod]
	public void SinShouldReturnCorrectResult() => Assert.AreEqual(0.909297427, ExecuteFunctionRounded(new Sin(), 9, 2));

	[TestMethod]
	public void SinhShouldReturnCorrectResult() => Assert.AreEqual(3.626860408d, ExecuteFunctionRounded(new Sinh(), 9, 2));

	[TestMethod]
	public void TanShouldReturnCorrectResult() => Assert.AreEqual(-2.185039863d, ExecuteFunctionRounded(new Tan(), 9, 2));

	[TestMethod]
	public void TanhShouldReturnCorrectResult() => Assert.AreEqual(0.96402758d, ExecuteFunctionRounded(new Tanh(), 9, 2));

	[TestMethod]
	public void AtanShouldReturnCorrectResult() => Assert.AreEqual(1.471127674d, ExecuteFunctionRounded(new Atan(), 9, 10));

	[TestMethod]
	public void Atan2ShouldReturnCorrectResult() => Assert.AreEqual(1.107148718d, ExecuteFunctionRounded(new Atan2(), 9, 1, 2));

	[TestMethod]
	public void AtanhShouldReturnCorrectResult() => Assert.AreEqual(0.1003d, ExecuteFunctionRounded(new Atanh(), 4, 0.1));

	[TestMethod]
	public void LogShouldReturnCorrectResult() => Assert.AreEqual(0.301029996d, ExecuteFunctionRounded(new Log(), 9, 2));

	[TestMethod]
	public void LogShouldReturnCorrectResultWithBase() => Assert.AreEqual(1d, ExecuteFunction(new Log(), 2, 2));

	[TestMethod]
	public void Log10ShouldReturnCorrectResult() => Assert.AreEqual(0.301029996d, ExecuteFunctionRounded(new Log10(), 9, 2));

	[TestMethod]
	public void LnShouldReturnCorrectResult() => Assert.AreEqual(1.60944d, ExecuteFunctionRounded(new Ln(), 5, 5));

	[TestMethod]
	public void SqrtPiShouldReturnCorrectResult() => Assert.AreEqual(2.506628275d, ExecuteFunctionRounded(new SqrtPi(), 9, 2));
}

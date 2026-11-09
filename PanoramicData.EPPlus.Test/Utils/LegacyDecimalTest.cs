using System.Globalization;
using OfficeOpenXml.Utils;

namespace PanoramicData.EPPlus.Test.Utils;

// Expected strings are (decimal)value as printed by the .NET 10.0.14 runtime.
[TestClass]
public class LegacyDecimalTest
{
	[TestMethod]
	[DataRow(4600808574327391672L, "0.396180555555556")]
	[DataRow(4608983858650965606L, "1.4")]
	[DataRow(4652517263873548186L, "1091.9503500182")]
	[DataRow(4862596447618666293L, "123456789012346000")]
	[DataRow(4593560419847042655L, "0.123456789012346")]
	[DataRow(4188260752279232537L, "0.0000000000000000000000000001")]
	[DataRow(4172965893781238599L, "0")]
	[DataRow(-9223372036854775808L, "0")]
	[DataRow(4607182418800017410L, "1")]
	[DataRow(5039524781320396022L, "79200000000000000000000000000")]
	[DataRow(-4183869992901924110L, "-79000000000000000000000000000")]
	public void FromDoubleMatchesNet10(long bits, string expected)
		=> Assert.AreEqual(expected, LegacyDecimal.FromDouble(BitConverter.Int64BitsToDouble(bits)).ToString(CultureInfo.InvariantCulture));

	[TestMethod]
	[DataRow(1036831949, "0.1")]
	[DataRow(1068708659, "1.4")]
	[DataRow(1842099665, "7900000000000000000000000000")]
	[DataRow(285050806, "0.0000000000000000000000000001")]
	public void FromSingleMatchesNet10(int bits, string expected)
		=> Assert.AreEqual(expected, LegacyDecimal.FromSingle(BitConverter.Int32BitsToSingle(bits)).ToString(CultureInfo.InvariantCulture));

	[TestMethod]
	[DataRow(5039527983027585024L)]
	[DataRow(9218868437227405312L)]
	[DataRow(9221120237041090560L)]
	public void FromDoubleOutOfRangeThrowsOverflow(long bits)
		=> Assert.ThrowsExactly<OverflowException>(() => LegacyDecimal.FromDouble(BitConverter.Int64BitsToDouble(bits)));

	[TestMethod]
	public void FromSingleOutOfRangeThrowsOverflow()
		=> Assert.ThrowsExactly<OverflowException>(() => LegacyDecimal.FromSingle(BitConverter.Int32BitsToSingle(2139095039)));

	[TestMethod]
	public void GetTypedCellValueRoundsDoubleTo15SignificantDigits()
		=> Assert.AreEqual(0.396180555555556m, ConvertUtil.GetTypedCellValue<decimal>(0.396180555555556));
}

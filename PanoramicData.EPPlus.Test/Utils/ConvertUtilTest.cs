using System.ComponentModel;
using System.Globalization;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;

namespace PanoramicData.EPPlus.Test.Utils;

[TestClass]
public class ConvertUtilTest
{
	private static void AssertParsesNumericString(object numericString, double expected)
	{
		Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out var result));
		Assert.AreEqual(expected, result);
	}

	private static void AssertParsesDateString(object dateString, DateTime expected)
	{
		Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out var result));
		Assert.AreEqual(expected, result);
	}

	[TestMethod]
	public void TryParseNumericString()
	{
		object numericString = null;
		double expected = 0;
		Assert.IsFalse(ConvertUtil.TryParseNumericString(numericString, out var result));
		Assert.AreEqual(expected, result);
		expected = 1442.0;
		numericString = expected.ToString("e", CultureInfo.CurrentCulture); // 1.442E+003
		AssertParsesNumericString(numericString, expected);
		numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // 1442
		AssertParsesNumericString(numericString, expected);
		numericString = expected.ToString("f2", CultureInfo.CurrentCulture); // 1442.00
		AssertParsesNumericString(numericString, expected);
		numericString = expected.ToString("n", CultureInfo.CurrentCulture); // 1,442.0
		AssertParsesNumericString(numericString, expected);
		expected = -0.00526;
		numericString = expected.ToString("e", CultureInfo.CurrentCulture); // -5.26E-003
		AssertParsesNumericString(numericString, expected);
		numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // -0
		AssertParsesNumericString(numericString, 0.0);
		numericString = expected.ToString("f3", CultureInfo.CurrentCulture); // -0.005
		AssertParsesNumericString(numericString, -0.005);
		numericString = expected.ToString("n6", CultureInfo.CurrentCulture); // -0.005260
		AssertParsesNumericString(numericString, expected);
	}

	[TestMethod]
	public void TryParseDateString()
	{
		object dateString = null;
		var expected = DateTime.MinValue;
		Assert.IsFalse(ConvertUtil.TryParseDateString(dateString, out var result));
		Assert.AreEqual(expected, result);
		expected = new DateTime(2013, 1, 15);
		dateString = expected.ToString("d", CultureInfo.CurrentCulture); // 1/15/2013
		AssertParsesDateString(dateString, expected);
		dateString = expected.ToString("D", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013
		AssertParsesDateString(dateString, expected);
		dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 12:00:00 AM
		AssertParsesDateString(dateString, expected);
		dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 12:00 AM
		AssertParsesDateString(dateString, expected);
		expected = new DateTime(2013, 1, 15, 15, 26, 32);
		dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 3:26:32 PM
		AssertParsesDateString(dateString, expected);
		dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 3:26 PM
		AssertParsesDateString(dateString, new DateTime(2013, 1, 15, 15, 26, 0));
	}

	[TestMethod]
	public void TextToInt()
	{
		var result = ConvertUtil.GetTypedCellValue<int>("204");

		Assert.AreEqual(204, result);
	}
	// This is just illustration of the bug in old implementation
	//[TestMethod]
	public static void TextToIntInOldImplementation()
	{
		var result = GetTypedValue<int>("204");

		Assert.AreEqual(204, result);
	}
	[TestMethod]
	public void DoubleToNullableInt()
	{
		var result = ConvertUtil.GetTypedCellValue<int?>(2D);

		Assert.AreEqual(2, result);
	}

	[TestMethod]
	public void StringToDecimal()
	{
		var decimalSign = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
		var result = ConvertUtil.GetTypedCellValue<decimal>($"1{decimalSign}4");

		Assert.AreEqual((decimal)1.4, result);
	}

	[TestMethod]
	public void EmptyStringToNullableDecimal()
	{
		var result = ConvertUtil.GetTypedCellValue<decimal?>("");
		Assert.IsNull(result);
	}

	[TestMethod]
	public void BlankStringToNullableDecimal()
	{
		var result = ConvertUtil.GetTypedCellValue<decimal?>(" ");

		Assert.IsNull(result);
	}

	[TestMethod]
	public void EmptyStringToDecimal() => Assert.ThrowsExactly<FormatException>(() => ConvertUtil.GetTypedCellValue<decimal>(""));

	[TestMethod]
	public void FloatingPointStringToInt() => Assert.ThrowsExactly<FormatException>(() => ConvertUtil.GetTypedCellValue<int>("1.4"));

	[TestMethod]
	public void IntToDateTime() => Assert.ThrowsExactly<InvalidCastException>(() => ConvertUtil.GetTypedCellValue<DateTime>(122));

	[TestMethod]
	public void IntToTimeSpan() => Assert.ThrowsExactly<InvalidCastException>(() => ConvertUtil.GetTypedCellValue<TimeSpan>(122));

	[TestMethod]
	public void IntStringToTimeSpan() => Assert.AreEqual(TimeSpan.FromDays(122), ConvertUtil.GetTypedCellValue<TimeSpan>("122"));

	[TestMethod]
	public void BoolToInt()
	{
		Assert.AreEqual(1, ConvertUtil.GetTypedCellValue<int>(true));
		Assert.AreEqual(0, ConvertUtil.GetTypedCellValue<int>(false));
	}

	[TestMethod]
	public void BoolToDecimal()
	{
		Assert.AreEqual(1m, ConvertUtil.GetTypedCellValue<decimal>(true));
		Assert.AreEqual(0m, ConvertUtil.GetTypedCellValue<decimal>(false));
	}

	[TestMethod]
	public void BoolToDouble()
	{
		Assert.AreEqual(1d, ConvertUtil.GetTypedCellValue<double>(true));
		Assert.AreEqual(0d, ConvertUtil.GetTypedCellValue<double>(false));
	}

	[TestMethod]
	public void BadTextToInt() => Assert.ThrowsExactly<FormatException>(() => ConvertUtil.GetTypedCellValue<int>("text1"));

	// previous implementation
	internal static T GetTypedValue<T>(object v)
	{
		if (v == null)
		{
			return default;
		}

		var fromType = v.GetType();
		var toType = typeof(T);

		var toType2 = GetNullableUnderlyingType(toType);
		if (fromType == toType || fromType == toType2)
		{
			return (T)v;
		}

		var cnv = TypeDescriptor.GetConverter(fromType);
		if (toType == typeof(DateTime) || toType2 == typeof(DateTime))    //Handle dates
		{
			return GetTypedDate<T>(v, fromType, cnv);
		}
		else if (toType == typeof(TimeSpan) || toType2 == typeof(TimeSpan))    //Handle timespan
		{
			return GetTypedTimeSpan<T>(v, fromType, cnv);
		}
		else
		{
			return GetTypedOther<T>(v, fromType, cnv, toType, toType2);
		}
	}

	private static Type GetNullableUnderlyingType(Type toType)
		=> TypeCompat.IsGenericType(toType) && toType.GetGenericTypeDefinition().Equals(typeof(Nullable<>))
			? Nullable.GetUnderlyingType(toType)
			: null;

	private static T GetTypedDate<T>(object v, Type fromType, TypeConverter cnv)
	{
		if (fromType == typeof(TimeSpan))
		{
			return (T)(object)new DateTime(((TimeSpan)v).Ticks);
		}
		else if (fromType == typeof(string))
		{
			return DateTime.TryParse(v.ToString(), out var dt) ? (T)(object)dt : default;

		}
		else
		{
			return cnv.CanConvertTo(typeof(double)) ? (T)(object)DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))) : default;
		}
	}

	private static T GetTypedTimeSpan<T>(object v, Type fromType, TypeConverter cnv)
	{
		if (fromType == typeof(DateTime))
		{
			return (T)(object)new TimeSpan(((DateTime)v).Ticks);
		}
		else if (fromType == typeof(string))
		{
			return TimeSpan.TryParse(v.ToString(), out var ts) ? (T)(object)ts : default;
		}
		else
		{
			if (cnv.CanConvertTo(typeof(double)))
			{

				return (T)(object)new TimeSpan(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))).Ticks);
			}
			else
			{
				try
				{
					// Issue 14682 -- "GetValue<decimal>() won't convert strings"
					// As suggested, after all special cases, all .NET to do it's 
					// preferred conversion rather than simply returning the default
					return (T)Convert.ChangeType(v, typeof(T));
				}
				catch (Exception)
				{
					// This was the previous behaviour -- no conversion is available.
					return default;
				}
			}
		}
	}

	private static T GetTypedOther<T>(object v, Type fromType, TypeConverter cnv, Type toType, Type toType2)
	{
		if (cnv.CanConvertTo(toType))
		{
			return (T)cnv.ConvertTo(v, typeof(T));
		}
		else
		{
			if (toType2 != null)
			{
				toType = toType2;
				if (cnv.CanConvertTo(toType))
				{
					return (T)cnv.ConvertTo(v, toType); //Fixes issue 15377
				}
			}

			if (fromType == typeof(double) && toType == typeof(decimal))
			{
				return (T)(object)Convert.ToDecimal(v);
			}
			else
			{
				return fromType == typeof(decimal) && toType == typeof(double) ? (T)(object)Convert.ToDouble(v) : default;
			}
		}
	}
}

using System.Globalization;

namespace OfficeOpenXml.Utils;

/// <summary>
/// Converts double and float to decimal the way .NET 10 and earlier did: rounded to 15 (double) or 7 (float) significant digits.
/// </summary>
internal static class LegacyDecimal
{
	private const int MaxScale = 28;

	private static readonly double[] _powersOfTen = BuildPowersOfTen();

	internal static decimal FromDouble(double value)
	{
		if (double.IsNaN(value) || double.IsInfinity(value))
		{
			throw new OverflowException("Value was either too large or too small for a Decimal.");
		}

		var exponent = (int)((ulong)BitConverter.DoubleToInt64Bits(value) >> 52 & 0x7FF) - 1022;
		return Round(value, exponent, 15);
	}

	internal static decimal FromSingle(float value)
	{
		if (float.IsNaN(value) || float.IsInfinity(value))
		{
			throw new OverflowException("Value was either too large or too small for a Decimal.");
		}

		var exponent = (int)((uint)BitConverter.SingleToInt32Bits(value) >> 23 & 0xFF) - 126;
		return Round(value, exponent, 7);
	}

	private static double[] BuildPowersOfTen()
	{
		var powers = new double[81];
		for (var i = 0; i < powers.Length; i++)
		{
			powers[i] = double.Parse("1e" + i.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
		}

		return powers;
	}

	// Port of the runtime's pre-.NET 11 VarDecFromR8/VarDecFromR4, including its double-precision scaling.
	private static decimal Round(double value, int exponent, int digits)
	{
		if (exponent < -94)
		{
			return 0m;
		}

		if (exponent > 96)
		{
			throw new OverflowException("Value was either too large or too small for a Decimal.");
		}

		var negative = value < 0;
		var scaled = Math.Abs(value);
		var limit = _powersOfTen[digits];
		var power = digits - 1 - (exponent * 19728 >> 16);
		if (power >= 0)
		{
			power = Math.Min(power, MaxScale);
			scaled *= _powersOfTen[power];
		}
		else if (power != -1 || scaled >= limit)
		{
			scaled /= _powersOfTen[-power];
		}
		else
		{
			power = 0;
		}

		if (scaled < limit / 10 && power < MaxScale)
		{
			scaled *= 10;
			power++;
		}

		var mantissa = (ulong)(long)Math.Round(scaled);
		if (mantissa == 0)
		{
			return 0m;
		}

		if (power < 0)
		{
			var result = (decimal)mantissa;
			for (var i = power; i < 0; i++)
			{
				result *= 10;
			}

			return negative ? -result : result;
		}

		var removable = Math.Min(power, digits - 1);
		while (removable > 0 && mantissa % 10 == 0)
		{
			mantissa /= 10;
			power--;
			removable--;
		}

		return new decimal((int)(uint)mantissa, (int)(uint)(mantissa >> 32), 0, negative, (byte)power);
	}
}

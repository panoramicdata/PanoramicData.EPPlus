using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

internal static class SignificanceRounding
{
	internal static void ValidateNumberAndSign(double number, double sign, string functionName)
	{
		if (number > 0d && sign < 0)
		{
			var values = string.Format("num: {0}, sign: {1}", number, sign);
			throw new InvalidOperationException(functionName + " cannot handle a negative significance when the number is positive" + values);
		}
	}

	internal static double RoundFractional(double number, double significance, int extraSteps)
	{
		var floor = System.Math.Floor(number);
		var rest = number - floor;
		var nSign = (int)(rest / significance) + extraSteps;
		return floor + (nSign * significance);
	}
}

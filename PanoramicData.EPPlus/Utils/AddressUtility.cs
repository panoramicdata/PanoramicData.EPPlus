﻿using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils;

public static class AddressUtility
{
	public static string ParseEntireColumnSelections(string address)
	{
		var parsedAddress = address;
		var matches = Regex.Matches(address, "[A-Z]+:[A-Z]+");
		foreach (Match match in matches)
		{
			AddRowNumbersToEntireColumnRange(ref parsedAddress, match.Value);
		}

		return parsedAddress;
	}

	private static void AddRowNumbersToEntireColumnRange(ref string address, string range)
	{
		var parsedRange = string.Format("{0}{1}", range, ExcelPackage.MaxRows);
		var splitArr = parsedRange.Split([':']);
		address = address.Replace(range, string.Format("{0}1:{1}", splitArr[0], splitArr[1]));
	}
}

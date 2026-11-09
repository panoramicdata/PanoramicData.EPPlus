using System;

namespace OfficeOpenXml.Table;

/// <summary>
/// Name and style handling shared by tables and pivot tables.
/// </summary>
internal static class TableNameHelper
{
	internal static void RenameTable(ExcelWorksheet worksheet, string prevName, string newName)
	{
		if (worksheet.Tables._tableNames.TryGetValue(prevName, out var ix))
		{
			worksheet.Tables._tableNames.Remove(prevName);
			worksheet.Tables._tableNames.Add(newName, ix);
		}
	}

	// "None" clears the stored style name, so the name to store is returned with the style.
	internal static (TableStyles Style, string Name) ParseStyleName(string value, string builtInPrefix)
	{
		if (value.StartsWith(builtInPrefix))
		{
			try
			{
				return (Enum.Parse<TableStyles>(value[builtInPrefix.Length..], true), value);
			}
			catch
			{
				return (TableStyles.Custom, value);
			}
		}

		if (value == "None")
		{
			return (TableStyles.None, "");
		}

		return (TableStyles.Custom, value);
	}
}

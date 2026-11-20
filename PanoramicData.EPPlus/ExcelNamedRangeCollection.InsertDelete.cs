namespace OfficeOpenXml;

public partial class ExcelNamedRangeCollection
{
	internal void Insert(int rowFrom, int colFrom, int rows, int cols) => Insert(rowFrom, colFrom, rows, cols, n => true);

	internal void Insert(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
	{
		var namedRanges = _list.Where(filter);
		foreach (var namedRange in namedRanges)
		{
			InsertRows(rowFrom, rows, namedRange);
			InsertColumns(colFrom, cols, namedRange);
		}
	}
	internal void Delete(int rowFrom, int colFrom, int rows, int cols) => Delete(rowFrom, colFrom, rows, cols, n => true);
	internal void Delete(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
	{
		var namedRanges = _list.Where(filter);
		foreach (var namedRange in namedRanges)
		{
			ExcelAddressBase adr;
			if (cols > 0 && rowFrom == 0 && rows >= ExcelPackage.MaxRows)   //Issue 15554. Check
			{
				adr = namedRange.DeleteColumn(colFrom, cols);
			}
			else
			{
				adr = namedRange.DeleteRow(rowFrom, rows);
			}

			if (adr == null)
			{
				namedRange.Address = "#REF!";
			}
			else
			{
				namedRange.Address = adr.Address;
			}
		}
	}

	private static void InsertColumns(int colFrom, int cols, ExcelNamedRange namedRange)
	{
		if (colFrom > 0)
		{
			ApplyInsert(namedRange, rng => GetInsertedColumnsAddress(rng, colFrom, cols));
		}
	}

	private static string? GetInsertedColumnsAddress(ExcelAddressBase rng, int colFrom, int cols)
	{
		if (colFrom <= rng.Start.Column)
		{
			return ExcelCellBase.GetAddress(rng.Start.Row, rng.Start.Column + cols, rng.End.Row, rng.End.Column + cols);
		}

		if (colFrom <= rng.End.Column && rng.End.Column + cols < ExcelPackage.MaxColumns)
		{
			return ExcelCellBase.GetAddress(rng.Start.Row, rng.Start.Column, rng.End.Row, rng.End.Column + cols);
		}

		return null;
	}

	private static void ApplyInsert(ExcelNamedRange namedRange, Func<ExcelAddressBase, string?> getNewAddress)
	{
		if (namedRange.Addresses?.Count > 0) //Contains multiple ranges
		{
			foreach (var rng in namedRange.Addresses)
			{
				var newAddress = getNewAddress(rng);
				if (newAddress != null)
				{
					rng.Address = BuildNewAddress(namedRange, newAddress);
				}
			}

			namedRange.Address = namedRange.FullAddress;
		}
		else
		{
			var newAddress = getNewAddress(namedRange);
			if (newAddress != null)
			{
				namedRange.Address = BuildNewAddress(namedRange, newAddress);
			}
		}
	}

	private static string BuildNewAddress(ExcelNamedRange namedRange, string newAddress)
	{
		if (namedRange.FullAddress.Contains('!'))
		{
			var worksheet = namedRange.FullAddress.Split('!')[0];
			worksheet = worksheet.Trim('\'');
			newAddress = ExcelCellBase.GetFullAddress(worksheet, newAddress);
		}

		return newAddress;
	}

	private static void InsertRows(int rowFrom, int rows, ExcelNamedRange namedRange)
	{
		if (rows > 0)
		{
			ApplyInsert(namedRange, rng => GetInsertedRowsAddress(rng, rowFrom, rows));
		}
	}

	private static string? GetInsertedRowsAddress(ExcelAddressBase rng, int rowFrom, int rows)
	{
		if (rowFrom <= rng.Start.Row)
		{
			return ExcelCellBase.GetAddress(rng.Start.Row + rows, rng.Start.Column, rng.End.Row + rows, rng.End.Column);
		}

		if (rowFrom <= rng.End.Row && rng.End.Row + rows <= ExcelPackage.MaxRows)
		{
			return ExcelCellBase.GetAddress(rng.Start.Row, rng.Start.Column, rng.End.Row + rows, rng.End.Column);
		}

		return null;
	}
}

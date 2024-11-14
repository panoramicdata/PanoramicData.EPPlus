using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public struct ExcelDoubleCellValue : IComparable<ExcelDoubleCellValue>, IComparable
{
	public ExcelDoubleCellValue(double val)
	{
		Value = val;
		CellRow = default;
	}

	public ExcelDoubleCellValue(double val, int cellRow)
	{
		Value = val;
		CellRow = cellRow;
	}

	public int? CellRow;

	public double Value;

	public static implicit operator double(ExcelDoubleCellValue d)
	{
		return d.Value;
	}
	//  User-defined conversion from double to Digit
	public static implicit operator ExcelDoubleCellValue(double d)
	{
		return new ExcelDoubleCellValue(d);
	}

	public int CompareTo(ExcelDoubleCellValue other) => Value.CompareTo(other.Value);

	public int CompareTo(object obj) => obj is double ? Value.CompareTo((double)obj) : Value.CompareTo(((ExcelDoubleCellValue)obj).Value);

	public override bool Equals(object obj) => CompareTo(obj) == 0;
	public override int GetHashCode() => base.GetHashCode();
	public static bool operator ==(ExcelDoubleCellValue a, ExcelDoubleCellValue b)
	{
		return a.Value.CompareTo(b.Value) == 0;
	}

	public static bool operator ==(ExcelDoubleCellValue a, double b)
	{
		return a.Value.CompareTo(b) == 0;
	}

	public static bool operator !=(ExcelDoubleCellValue a, ExcelDoubleCellValue b)
	{
		return a.Value.CompareTo(b.Value) != 0;
	}

	public static bool operator !=(ExcelDoubleCellValue a, double b)
	{
		return a.Value.CompareTo(b) != 0;
	}
}

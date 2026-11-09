using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	private static readonly _setValue _setStyleIdDelegate = Set_StyleID;
	private static readonly _setValue _setValueDelegate = Set_Value;
	private static readonly _setValue _setHyperLinkDelegate = Set_HyperLink;
	private static readonly _setValue _setIsRichTextDelegate = Set_IsRichText;
	private static readonly _setValue _setExistsCommentDelegate = Exists_Comment;
	private static readonly _setValue _setCommentDelegate = Set_Comment;

	private static void Set_StyleID(ExcelRangeBase range, object value, int row, int col) => range._worksheet.SetStyleInner(row, col, (int)value);
	private static void Set_Value(ExcelRangeBase range, object value, int row, int col)
	{
		var sfi = range._worksheet._formulas.GetValue(row, col);
		if (sfi is int)
		{
			range.SplitFormulas(range._worksheet.Cells[row, col]);
		}

		if (sfi != null)
		{
			range._worksheet._formulas.SetValue(row, col, string.Empty);
		}

		range._worksheet.SetValueInner(row, col, value);
	}
	private static void Set_Formula(ExcelRangeBase range, object value, int row, int col)
	{
		var f = range._worksheet._formulas.GetValue(row, col);
		if (f is int v && v >= 0)
		{
			range.SplitFormulas(range._worksheet.Cells[row, col]);
		}

		var formula = value == null ? string.Empty : value.ToString();
		if (formula == string.Empty)
		{
			range._worksheet._formulas.SetValue(row, col, string.Empty);
		}
		else
		{
			range._worksheet._formulas.SetValue(row, col, formula);
			range._worksheet.SetValueInner(row, col, null);
		}
	}
	/// <summary>
	/// Handles shared formulas
	/// </summary>
	/// <param name="range">The range</param>
	/// <param name="value">The  formula</param>
	/// <param name="address">The address of the formula</param>
	/// <param name="IsArray">If the forumla is an array formula.</param>
	private static void Set_SharedFormula(ExcelRangeBase range, string value, ExcelAddress address, bool IsArray)
	{
		if (range._fromRow == 1 && range._fromCol == 1 && range._toRow == ExcelPackage.MaxRows && range._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
		{
			throw new InvalidOperationException("Can't set a formula for the entire worksheet");
		}
		else if (address.Start.Row == address.End.Row && address.Start.Column == address.End.Column && !IsArray)             //is it really a shared formula? Arrayformulas can be one cell only
		{
			//Nope, single cell. Set the formula
			Set_Formula(range, value, address.Start.Row, address.Start.Column);
			return;
		}

		range.CheckAndSplitSharedFormula(address);
		ExcelWorksheet.Formulas f = new(SourceCodeTokenizer.Default)
		{
			Formula = value,
			Index = range._worksheet.GetMaxShareFunctionIndex(IsArray),
			Address = address.FirstAddress,
			StartCol = address.Start.Column,
			StartRow = address.Start.Row,
			IsArray = IsArray
		};

		range._worksheet._sharedFormulas.Add(f.Index, f);

		for (var col = address.Start.Column; col <= address.End.Column; col++)
		{
			for (var row = address.Start.Row; row <= address.End.Row; row++)
			{
				range._worksheet._formulas.SetValue(row, col, f.Index);
				range._worksheet._flags.SetFlagValue(row, col, true, CellFlags.ArrayFormula);
				range._worksheet.SetValueInner(row, col, null);
			}
		}
	}
	private static void Set_HyperLink(ExcelRangeBase range, object value, int row, int col)
	{
		if (value is Uri uri)
		{
			range._worksheet._hyperLinks.SetValue(row, col, uri);

			if (value is ExcelHyperLink link)
			{
				range._worksheet.SetValueInner(row, col, link.Display);
			}
			else
			{
				var v = range._worksheet.GetValueInner(row, col);
				if (v == null || v.ToString() == "")
				{
					range._worksheet.SetValueInner(row, col, uri.OriginalString);
				}
			}
		}
		else
		{
			range._worksheet._hyperLinks.SetValue(row, col, (Uri)null);
			range._worksheet.SetValueInner(row, col, (Uri)null);
		}
	}
	private static void Set_IsRichText(ExcelRangeBase range, object value, int row, int col) => range._worksheet._flags.SetFlagValue(row, col, (bool)value, CellFlags.RichText);
	private static void Exists_Comment(ExcelRangeBase range, object value, int row, int col)
	{
		if (range._worksheet._commentsStore.Exists(row, col))
		{
			throw new InvalidOperationException(string.Format("Cell {0} already contain a comment.", new ExcelCellAddress(row, col).Address));
		}

	}
	private static void Set_Comment(ExcelRangeBase range, object value, int row, int col)
	{
		var v = (string[])value;
		range._worksheet.Comments.Add(new ExcelRangeBase(range._worksheet, GetAddress(range._fromRow, range._fromCol)), v[0], v[1]);
	}
}

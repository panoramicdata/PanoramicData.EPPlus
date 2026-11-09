using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Gets or sets a formula for a range.
	/// </summary>
	public string Formula
	{
		get
		{
			if (IsName)
			{
				return _worksheet == null ? _workbook._names[_address].NameFormula : _worksheet.Names[_address].NameFormula;
			}
			else
			{
				return _worksheet.GetFormula(_fromRow, _fromCol);
			}
		}
		set
		{
			if (IsName)
			{
				if (_worksheet == null)
				{
					_workbook._names[_address].NameFormula = value;
				}
				else
				{
					_worksheet.Names[_address].NameFormula = value;
				}
			}
			else
			{
				if (value == null || value.Trim() == "")
				{
					//Set the cells to null
					Value = null;
				}
				else if (_fromRow == _toRow && _fromCol == _toCol)
				{
					Set_Formula(this, value, _fromRow, _fromCol);
				}
				else
				{
					Set_SharedFormula(this, value, this, false);
					if (Addresses != null)
					{
						foreach (var address in Addresses)
						{
							Set_SharedFormula(this, value, address, false);
						}
					}
				}
			}
		}
	}
	/// <summary>
	/// Gets or Set a formula in R1C1 format.
	/// </summary>
	public string FormulaR1C1
	{
		get
		{
			IsRangeValid("FormulaR1C1");
			return _worksheet.GetFormulaR1C1(_fromRow, _fromCol);
		}
		set
		{
			IsRangeValid("FormulaR1C1");
			if (value.Length > 0 && value[0] == '=')
			{
				value = value[1..]; // remove any starting equalsign.
			}

			if (value == null || value.Trim() == "")
			{
				//Set the cells to null
				_worksheet.Cells[TranslateFromR1C1(value, _fromRow, _fromCol)].Value = null;
			}
			else if (Addresses == null)
			{
				Set_SharedFormula(this, TranslateFromR1C1(value, _fromRow, _fromCol), this, false);
			}
			else
			{
				Set_SharedFormula(this, TranslateFromR1C1(value, _fromRow, _fromCol), new ExcelAddress(WorkSheet, FirstAddress), false);
				foreach (var address in Addresses)
				{
					Set_SharedFormula(this, TranslateFromR1C1(value, address.Start.Row, address.Start.Column), address, false);
				}
			}
		}
	}
	/// <summary>
	/// Is the range a part of an Arrayformula
	/// </summary>
	public bool IsArrayFormula
	{
		get
		{
			IsRangeValid("arrayformulas");
			return _worksheet._flags.GetFlagValue(_fromRow, _fromCol, CellFlags.ArrayFormula);
		}
	}
	/// <summary>
	/// Creates an array-formula.
	/// </summary>
	/// <param name="ArrayFormula">The formula</param>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public void CreateArrayFormula(string ArrayFormula)
	{
		if (Addresses != null)
		{
			throw new Exception("An Arrayformula can not have more than one address");
		}

		Set_SharedFormula(this, ArrayFormula, this, true);
	}
}

using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System.Text;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	private void LoadFormulaElement(XmlReader xr, ExcelAddressBase address)
	{
		var t = xr.GetAttribute("t");
		if (t == null)
		{
			_formulas.SetValue(address._fromRow, address._fromCol, xr.ReadElementContentAsString());
			SetValueInner(address._fromRow, address._fromCol, null);
		}
		else if (t == "shared")
		{
			LoadSharedFormula(xr, address);
		}
		else if (t == "array") //TODO: Array functions are not support yet. Read the formula for the start cell only.
		{
			LoadArrayFormula(xr, address);
		}
		else // ??? some other type
		{
			xr.Read();  //Something is wrong in the sheet, read next
		}

	}

	private void LoadSharedFormula(XmlReader xr, ExcelAddressBase address)
	{
		var si = xr.GetAttribute("si");
		if (si != null)
		{
			var sfIndex = int.Parse(si);
			_formulas.SetValue(address._fromRow, address._fromCol, sfIndex);
			SetValueInner(address._fromRow, address._fromCol, null);
			var fAddress = xr.GetAttribute("ref");
			var formula = ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString());
			if (formula != "")
			{
				_sharedFormulas.Add(sfIndex, new Formulas(SourceCodeTokenizer.Default) { Index = sfIndex, Formula = formula, Address = fAddress, StartRow = address._fromRow, StartCol = address._fromCol });
			}
		}
		else
		{
			xr.Read();  //Something is wrong in the sheet, read next
		}
	}

	private void LoadArrayFormula(XmlReader xr, ExcelAddressBase address)
	{
		var aAddress = xr.GetAttribute("ref");
		var formula = xr.ReadElementContentAsString();
		var afIndex = GetMaxShareFunctionIndex(true);
		_formulas.SetValue(address._fromRow, address._fromCol, afIndex);
		SetValueInner(address._fromRow, address._fromCol, null);
		_sharedFormulas.Add(afIndex, new Formulas(SourceCodeTokenizer.Default) { Index = afIndex, Formula = formula, Address = aAddress, StartRow = address._fromRow, StartCol = address._fromCol, IsArray = true });
		_flags.SetFlagValue(address._fromRow, address._fromCol, true, CellFlags.ArrayFormula);
	}

	private void LoadInlineString(XmlReader xr, ExcelAddressBase address)
	{
		xr.Read();
		if (xr.LocalName == "t")
		{
			SetValueInner(address._fromRow, address._fromCol, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
		}
		else
		{
			if (xr.LocalName == "r")
			{
				var rXml = new StringBuilder(xr.ReadOuterXml());
				while (xr.LocalName == "r")
				{
					rXml.Append(xr.ReadOuterXml());
				}

				SetValueInner(address._fromRow, address._fromCol, rXml.ToString());
			}
			else
			{
				SetValueInner(address._fromRow, address._fromCol, xr.ReadOuterXml());
			}
			_flags.SetFlagValue(address._fromRow, address._fromCol, true, CellFlags.RichText);
		}
	}
}

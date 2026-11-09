using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Load cells
	/// </summary>
	/// <param name="xr">The reader</param>
	private void LoadCells(XmlReader xr)
	{
		ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
		ExcelAddressBase address = null;
		var type = "";
		var style = 0;
		var row = 0;
		var col = 0;
		xr.Read();

		while (!xr.EOF)
		{
			while (xr.NodeType == XmlNodeType.EndElement)
			{
				xr.Read();
				continue;
			}

			if (xr.LocalName == "row")
			{
				col = 0;
				var r = xr.GetAttribute("r");
				if (r == null)
				{
					row++;
				}
				else
				{
					row = Convert.ToInt32(r);
				}

				if (DoAddRow(xr))
				{
					SetValueInner(row, 0, AddRow(xr));
					if (xr.GetAttribute("s") != null)
					{
						SetStyleInner(row, 0, int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture));
					}
				}

				xr.Read();
			}
			else if (xr.LocalName == "c")
			{
				var r = xr.GetAttribute("r");
				if (r == null)
				{
					//Handle cells with no reference
					col++;
					address = new ExcelAddressBase(row, col, row, col);
				}
				else
				{
					address = new ExcelAddressBase(r);
					col = address._fromCol;
				}


				//Datetype
				if (xr.GetAttribute("t") != null)
				{
					type = xr.GetAttribute("t");
				}
				else
				{
					type = "";
				}
				//Style
				if (xr.GetAttribute("s") != null)
				{
					style = int.Parse(xr.GetAttribute("s"));
					SetStyleInner(address._fromRow, address._fromCol, style);
					//SetValueInner(address._fromRow, address._fromCol, null); //TODO:Better Performance ??
				}
				else
				{
					style = 0;
				}

				xr.Read();
			}
			else if (xr.LocalName == "v")
			{
				SetValueFromXml(xr, type, style, address._fromRow, address._fromCol);

				xr.Read();
			}
			else if (xr.LocalName == "f")
			{
				var t = xr.GetAttribute("t");
				if (t == null)
				{
					_formulas.SetValue(address._fromRow, address._fromCol, xr.ReadElementContentAsString());
					SetValueInner(address._fromRow, address._fromCol, null);
				}
				else if (t == "shared")
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
				else if (t == "array") //TODO: Array functions are not support yet. Read the formula for the start cell only.
				{
					var aAddress = xr.GetAttribute("ref");
					var formula = xr.ReadElementContentAsString();
					var afIndex = GetMaxShareFunctionIndex(true);
					_formulas.SetValue(address._fromRow, address._fromCol, afIndex);
					SetValueInner(address._fromRow, address._fromCol, null);
					_sharedFormulas.Add(afIndex, new Formulas(SourceCodeTokenizer.Default) { Index = afIndex, Formula = formula, Address = aAddress, StartRow = address._fromRow, StartCol = address._fromCol, IsArray = true });
					_flags.SetFlagValue(address._fromRow, address._fromCol, true, CellFlags.ArrayFormula);
				}
				else // ??? some other type
				{
					xr.Read();  //Something is wrong in the sheet, read next
				}

			}
			else if (xr.LocalName == "is")   //Inline string
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
			else
			{
				break;
			}
		}

	}

	private static bool DoAddRow(XmlReader xr)
	{
		var c = xr.GetAttribute("r") == null ? 0 : 1;
		if (xr.GetAttribute("spans") != null)
		{
			c++;
		}

		return xr.AttributeCount > c;
	}
}

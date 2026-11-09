using System.Globalization;
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
			}

			if (xr.LocalName == "row")
			{
				col = 0;
				LoadRowElement(xr, ref row);
			}
			else if (xr.LocalName == "c")
			{
				address = LoadCellElement(xr, row, ref col, out type, out style);
			}
			else if (xr.LocalName == "v")
			{
				SetValueFromXml(xr, type, style, address._fromRow, address._fromCol);

				xr.Read();
			}
			else if (xr.LocalName == "f")
			{
				LoadFormulaElement(xr, address);
			}
			else if (xr.LocalName == "is")   //Inline string
			{
				LoadInlineString(xr, address);
			}
			else
			{
				break;
			}
		}

	}

	private void LoadRowElement(XmlReader xr, ref int row)
	{
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

	private ExcelAddressBase LoadCellElement(XmlReader xr, int row, ref int col, out string type, out int style)
	{
		ExcelAddressBase address;
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
		return address;
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

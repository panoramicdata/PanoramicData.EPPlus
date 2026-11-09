using OfficeOpenXml.Utils;
using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public sealed partial class ExcelWorkbook
{
	/// <summary>
	/// Max font width for the workbook
	/// <remarks>Taken from a built-in table of font metrics. For a font that is not in the table, set this value manually.</remarks>
	/// </summary>
	public decimal MaxFontWidth
	{
		get
		{
			var ix = Styles.NamedStyles.FindIndexByID("Normal");
			if (ix >= 0)
			{
				if (_standardFontWidth == decimal.MinValue || _fontID != Styles.NamedStyles[ix].Style.Font.Id)
				{
					var font = Styles.NamedStyles[ix].Style.Font;
					try
					{
						_standardFontWidth = GetWidthPixels(font.Name, font.Size);
						_fontID = Styles.NamedStyles[ix].Style.Font.Id;
					}
					catch   //Error, Font missing and Calibri removed in dictionary
					{
						_standardFontWidth = (int)(font.Size * (2D / 3D)); //Aprox for Calibri.
					}
				}
			}
			else
			{
				_standardFontWidth = 7; //Calibri 11
			}

			return _standardFontWidth;
		}
		set
		{
			_standardFontWidth = value;
		}
	}

	[SuppressMessage("Major Bug", "S1244:Floating point numbers should not be tested for equality", Justification = "min and max are both dictionary keys, so exact equality is the intended test.")]
	internal static decimal GetWidthPixels(string fontName, float fontSize)
	{
		Dictionary<float, FontSizeInfo> font;
		if (FontSize.FontHeights.TryGetValue(fontName, out var value))
		{
			font = value;
		}
		else
		{
			font = FontSize.FontHeights["Calibri"];
		}

		if (font.TryGetValue(fontSize, out var value2))
		{
			return LegacyDecimal.FromSingle(value2.Width);
		}
		else
		{
			float min = -1, max = 500;
			foreach (var size in font)
			{
				if (min < size.Key && size.Key < fontSize)
				{
					min = size.Key;
				}

				if (max > size.Key && size.Key > fontSize)
				{
					max = size.Key;
				}
			}

			return min == max
				? LegacyDecimal.FromSingle(font[min].Width)
				: LegacyDecimal.FromSingle(font[min].Height + (font[max].Height - font[min].Height) * ((fontSize - min) / (max - min)));
		}
	}
}

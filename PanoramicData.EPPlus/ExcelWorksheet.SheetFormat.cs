using OfficeOpenXml.Style.XmlAccess;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Get/set the default height of all rows in the worksheet
	/// </summary>
	public double DefaultRowHeight
	{
		get
		{
			CheckSheetType();
			var defaultRowHeight = GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight");
			if (double.IsNaN(defaultRowHeight) || CustomHeight == false)
			{
				defaultRowHeight = GetRowHeightFromNormalStyle();
			}

			return defaultRowHeight;
		}
		set
		{
			CheckSheetType();
			if (double.IsNaN(value))
			{
				DeleteNode("d:sheetFormatPr/@defaultRowHeight");
			}
			else
			{
				SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", value.ToString(CultureInfo.InvariantCulture));
				_ = GetRowHeightFromNormalStyle();
				CustomHeight = true;
			}
		}
	}

	private double GetRowHeightFromNormalStyle()
	{
		var ix = Workbook.Styles.NamedStyles.FindIndexByID("Normal");
		if (ix >= 0)
		{
			var f = Workbook.Styles.NamedStyles[ix].Style.Font;
			return ExcelFontXml.GetFontHeight(f.Name, f.Size) * 0.75;
		}
		else
		{
			return 15;   //Default Calibri 11
		}
	}

	/// <summary>
	/// 'True' if defaultRowHeight value has been manually set, or is different from the default value.
	/// Is automaticlly set to 'True' when assigning the DefaultRowHeight property
	/// </summary>
	public bool CustomHeight
	{
		get
		{
			return GetXmlNodeBool("d:sheetFormatPr/@customHeight");
		}
		set
		{
			SetXmlNodeBool("d:sheetFormatPr/@customHeight", value);
		}
	}
	/// <summary>
	/// Get/set the default width of all columns in the worksheet
	/// </summary>
	[SuppressMessage("Major Bug", "S1244:Floating point numbers should not be tested for equality", Justification = "Behaviour-preserving: a tolerance in place of the exact modulo test would change the computed default width.")]
	public double DefaultColWidth
	{
		get
		{
			CheckSheetType();
			var ret = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
			if (double.IsNaN(ret))
			{
				var mfw = Convert.ToDouble(Workbook.MaxFontWidth);
				var widthPx = mfw * 7;
				var margin = Math.Truncate(mfw / 4 + 0.999) * 2 + 1;
				if (margin < 5)
				{
					margin = 5;
				}

				while (Math.Truncate((widthPx - margin) / mfw * 100 + 0.5) / 100 < 8)
				{
					widthPx++;
				}

				widthPx = widthPx % 8 == 0 ? widthPx : 8 - widthPx % 8 + widthPx;
				var width = Math.Truncate((widthPx - margin) / mfw * 100 + 0.5) / 100;
				return Math.Truncate((width * mfw + margin) / mfw * 256) / 256;
			}

			return ret;
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString("d:sheetFormatPr/@defaultColWidth", value.ToString(CultureInfo.InvariantCulture));

			if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight")))
			{
				SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", GetRowHeightFromNormalStyle().ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	/** <outlinePr applyStyles="1" summaryBelow="0" summaryRight="0" /> **/
	private const string OutLineSummaryBelowPath = "d:sheetPr/d:outlinePr/@summaryBelow";

	/// <summary>
	/// Summary rows below details 
	/// </summary>
	public bool OutLineSummaryBelow
	{
		get
		{
			CheckSheetType();
			return GetXmlNodeBool(OutLineSummaryBelowPath);
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString(OutLineSummaryBelowPath, value ? "1" : "0");
		}
	}

	const string OutLineSummaryRightPath = "d:sheetPr/d:outlinePr/@summaryRight";

	/// <summary>
	/// Summary rows to right of details
	/// </summary>
	public bool OutLineSummaryRight
	{
		get
		{
			CheckSheetType();
			return GetXmlNodeBool(OutLineSummaryRightPath);
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString(OutLineSummaryRightPath, value ? "1" : "0");
		}
	}

	private const string OutLineApplyStylePath = "d:sheetPr/d:outlinePr/@applyStyles";

	/// <summary>
	/// Automatic styles
	/// </summary>
	public bool OutLineApplyStyle
	{
		get
		{
			CheckSheetType();
			return GetXmlNodeBool(OutLineApplyStylePath);
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString(OutLineApplyStylePath, value ? "1" : "0");
		}
	}

	private const string TabColorPath = "d:sheetPr/d:tabColor/@rgb";

	/// <summary>
	/// Color of the sheet tab
	/// </summary>
	public Color TabColor
	{
		get
		{
			var col = GetXmlNodeString(TabColorPath);
			return col == "" ? Color.Empty : Color.FromArgb(int.Parse(col, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			SetXmlNodeString(TabColorPath, value.ToArgb().ToString("X"));
		}
	}

	private const string CodeModuleNamePath = "d:sheetPr/@codeName";

	internal string CodeModuleName
	{
		get
		{
			return GetXmlNodeString(CodeModuleNamePath);
		}
		set
		{
			SetXmlNodeString(CodeModuleNamePath, value);
		}
	}

	internal void CodeNameChange(string value) => CodeModuleName = value;

	public VBA.ExcelVBAModule CodeModule => _package.Workbook.VbaProject != null ? _package.Workbook.VbaProject.Modules[CodeModuleName] : null;
}

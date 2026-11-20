namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Address for autofilter
	/// <seealso cref="ExcelRangeBase.AutoFilter" />        
	/// </summary>
	public ExcelAddressBase? AutoFilterAddress
	{
		get
		{
			CheckSheetType();
			var address = GetXmlNodeString("d:autoFilter/@ref");
			return address == "" ? null : new ExcelAddressBase(address);
		}
		internal set
		{
			CheckSheetType();
			if (value == null)
			{
				DeleteAllNode("d:autoFilter/@ref");
			}
			else
			{
				SetXmlNodeString("d:autoFilter/@ref", value.Address);
			}
		}
	}

	internal void CheckSheetType()
	{
		if (this is ExcelChartsheet)
		{
			throw new NotSupportedException("This property or method is not supported for a Chartsheet");
		}
	}

	/// <summary>
	/// Returns a ExcelWorksheetView object that allows you to set the view state properties of the worksheet
	/// </summary>
	public ExcelWorksheetView View
	{
		get
		{
			if (_sheetView == null)
			{
				var node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
				if (node == null)
				{
					CreateNode("d:sheetViews/d:sheetView");     //this one shouls always exist. but check anyway
					node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
				}

				_sheetView = new ExcelWorksheetView(NameSpaceManager, node, this);
			}

			return _sheetView;
		}
	}

	/// <summary>
	/// The worksheet's display name as it appears on the tab
	/// </summary>
	public string Name
	{
		get { return _name; }
		set
		{
			if (value == _name)
			{
				return;
			}

			value = ExcelWorksheets.ValidateFixSheetName(value);
			foreach (var ws in Workbook.Worksheets)
			{
				if (ws.PositionID != PositionID && ws.Name.Equals(value, StringComparison.OrdinalIgnoreCase))
				{
					throw new ArgumentException("Worksheet name must be unique");
				}
			}

			_package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@name", _sheetID), value);
			ChangeNames(value);

			_name = value;
		}
	}

	private void ChangeNames(string value)
	{
		// Renames name in this Worksheet
		foreach (var n in Workbook.Names)
		{
			if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue == null)
			{
				n.ChangeWorksheet(_name, value);
			}
		}

		foreach (var ws in Workbook.Worksheets)
		{
			if (ws is not ExcelChartsheet)
			{
				foreach (var n in ws.Names)
				{
					if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue == null)
					{
						n.ChangeWorksheet(_name, value);
					}
				}

				ws.UpdateCrossSheetReferenceNames(_name, value);
			}
		}
	}
	internal ExcelNamedRangeCollection _names;
	/// <summary>
	/// Provides access to named ranges
	/// </summary>
	public ExcelNamedRangeCollection Names
	{
		get
		{
			CheckSheetType();
			return _names;
		}
	}
	/// <summary>
	/// Indicates if the worksheet is hidden in the workbook
	/// </summary>
	public eWorkSheetHidden Hidden
	{
		get
		{
			var state = _package.Workbook.GetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
			if (state == "hidden")
			{
				return eWorkSheetHidden.Hidden;
			}
			else if (state == "veryHidden")
			{
				return eWorkSheetHidden.VeryHidden;
			}

			return eWorkSheetHidden.Visible;
		}
		set
		{
			if (value == eWorkSheetHidden.Visible)
			{
				_package.Workbook.DeleteNode(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
			}
			else
			{
				string v;
				v = value.ToString();
				v = v[..1].ToLowerInvariant() + v[1..];
				_package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID), v);
			}
		}
	}
}

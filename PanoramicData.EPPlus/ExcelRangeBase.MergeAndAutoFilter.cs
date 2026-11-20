namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// If the cells in the range are merged.
	/// </summary>
	public bool Merge
	{
		get
		{
			IsRangeValid("merging");
			for (var col = _fromCol; col <= _toCol; col++)
			{
				for (var row = _fromRow; row <= _toRow; row++)
				{
					if (_worksheet.MergedCells[row, col] == null)
					{
						return false;
					}
				}
			}

			return true;
		}
		set
		{
			IsRangeValid("merging");
			_worksheet.MergedCells.Clear(this);
			if (value)
			{
				_worksheet.MergedCells.Add(new ExcelAddressBase(FirstAddress), true);
				if (Addresses != null)
				{
					foreach (var address in Addresses)
					{
						_worksheet.MergedCells.Clear(address); //Fixes issue 15482
						_worksheet.MergedCells.Add(address, true);
					}
				}
			}
			else
			{
				if (Addresses != null)
				{
					foreach (var address in Addresses)
					{
						_worksheet.MergedCells.Clear(address);
					}
				}

			}
		}
	}
	/// <summary>
	/// Set an autofilter for the range
	/// </summary>
	public bool AutoFilter
	{
		get
		{
			IsRangeValid("autofilter");
			var address = _worksheet.AutoFilterAddress;
			if (address == null)
			{
				return false;
			}

			return _fromRow >= address.Start.Row
					&&
					_toRow <= address.End.Row
					&&
					_fromCol >= address.Start.Column
					&&
					_toCol <= address.End.Column;
		}
		set
		{
			IsRangeValid("autofilter");
			if (_worksheet.AutoFilterAddress != null)
			{
				var c = Collide(_worksheet.AutoFilterAddress);
				if (value == false && (c == eAddressCollition.Partly || c == eAddressCollition.No))
				{
					throw new InvalidOperationException("Can't remote Autofilter. Current autofilter does not match selected range.");
				}
			}

			if (_worksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
			{
				_worksheet.Names.Remove("_xlnm._FilterDatabase");
			}

			if (value)
			{
				_worksheet.AutoFilterAddress = this;
				var result = _worksheet.Names.Add("_xlnm._FilterDatabase", this);
				result.IsNameHidden = true;
			}
			else
			{
				_worksheet.AutoFilterAddress = null;
			}
		}
	}
}

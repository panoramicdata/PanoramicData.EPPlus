using System.Text;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	protected ExcelCellAddress? _start;
	/// <summary>
	/// Gets the row and column of the top left cell.
	/// </summary>
	/// <value>The start row column.</value>
	public ExcelCellAddress Start
	{
		get
		{
			_start ??= new ExcelCellAddress(_fromRow, _fromCol);

			return _start;
		}
	}
	protected ExcelCellAddress? _end;
	/// <summary>
	/// Gets the row and column of the bottom right cell.
	/// </summary>
	/// <value>The end row column.</value>
	public ExcelCellAddress End
	{
		get
		{
			_end ??= new ExcelCellAddress(_toRow, _toCol);

			return _end;
		}
	}
	protected ExcelTableAddress? _table;
	public ExcelTableAddress? Table => _table;

	/// <summary>
	/// The address for the range
	/// </summary>
	public virtual string Address => _address;
	internal string FullAddress
	{
		get
		{
			if (_addresses == null)
			{
				return GetAddress();
			}

			var a = new StringBuilder();
			foreach (var sa in _addresses)
			{
				a.Append(',').Append(sa.GetAddress());
			}

			return a.ToString();
		}
	}
	/// <summary>
	/// If the address is a defined name
	/// </summary>
	public bool IsName => _fromRow < 0;
	/// <summary>
	/// Returns the address text
	/// </summary>
	/// <returns></returns>
	public override string ToString() => _address;
	string _firstAddress;
	/// <summary>
	/// returns the first address if the address is a multi address.
	/// A1:A2,B1:B2 returns A1:A2
	/// </summary>
	internal string FirstAddress => string.IsNullOrEmpty(_firstAddress) ? _address : _firstAddress;
	internal string AddressSpaceSeparated => _address.Replace(',', ' '); //Conditional formatting and a few other places use space as separator for mulit addresses.
	/// <summary>
	/// Validate the address
	/// </summary>
	protected void Validate()
	{
		if (_fromRow > _toRow || _fromCol > _toCol)
		{
			throw new ArgumentOutOfRangeException("Start cell Address must be less or equal to End cell address");
		}
	}
	internal string WorkSheet => _ws;
	internal protected List<ExcelAddress>? _addresses;
	internal virtual List<ExcelAddress>? Addresses => _addresses;

	public int Rows => _toRow - _fromRow + 1;
	public int Columns => _toCol - _fromCol + 1;

	internal bool IsSingleCell => (_fromRow == _toRow && _fromCol == _toCol);
}

using System.Text;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private void SetWbWs(string address)
	{
		int pos;
		if (address[0] == '[')
		{
			pos = address.IndexOf(']');
			_wb = address[1..pos];
			_ws = address[(pos + 1)..];
		}
		else
		{
			_wb = "";
			_ws = address;
		}

		if (TrySetQuotedWorksheet())
		{
			return;
		}

		pos = _ws.IndexOf('!');

		if (pos == 0)
		{
			_address = _ws[1..];
			_ws = _wb;
			_wb = "";
		}
		else if (pos > -1)
		{
			_address = _ws[(pos + 1)..];
			_ws = _ws[..pos];
		}
		else
		{
			_address = address;
		}
	}

	private bool TrySetQuotedWorksheet()
	{
		if (!_ws.StartsWith('\''))
		{
			return false;
		}

		var pos = _ws.IndexOf('\'', 1);
		while (pos > 0 && pos + 1 < _ws.Length && _ws[pos + 1] == '\'')
		{
			_ws = _ws[..pos] + _ws[(pos + 1)..];
			pos = _ws.IndexOf('\'', pos + 1);
		}

		if (pos > 0)
		{
			_address = _ws[(pos + 2)..];
			_ws = _ws[1..pos];
			return true;
		}

		return false;
	}
	internal void ChangeWorksheet(string wsName, string newWs)
	{
		if (_ws == wsName)
		{
			_ws = newWs;
		}

		var fullAddress = new StringBuilder(GetAddress());

		if (Addresses != null)
		{
			foreach (var a in Addresses)
			{
				if (a._ws == wsName)
				{
					a._ws = newWs;
					fullAddress.Append(',').Append(a.GetAddress());
				}
				else
				{
					fullAddress.Append(',').Append(a._address);
				}
			}
		}

		_address = fullAddress.ToString();
	}

	private string GetAddress()
	{
		var adr = "";
		if (!string.IsNullOrEmpty(_wb))
		{
			adr = "[" + _wb + "]";
		}

		if (!string.IsNullOrEmpty(_ws))
		{
			adr += string.Format("'{0}'!", _ws);
		}

		if (IsName)
		{
			adr += GetAddress(_fromRow, _fromCol, _toRow, _toCol);
		}
		else
		{
			adr += GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}

		return adr;
	}
}

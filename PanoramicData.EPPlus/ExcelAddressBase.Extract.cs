using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeOpenXml;

public partial class ExcelAddressBase
{
	private void ExtractAddress(string fullAddress)
	{
		var brackPos = new Stack<int>();
		var bracketParts = new List<string>();
		string first = "", second = "";
		bool isText = false, hasSheet = false;
		try
		{
			if (fullAddress == "#REF!")
			{
				SetAddress(ref fullAddress, ref second, ref hasSheet, false);
				return;
			}
			else if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "!"))
			{
				// invalid address!
				return;
			}

			var isMulti = false;
			for (var i = 0; i < fullAddress.Length; i++)
			{
				var c = fullAddress[i];
				if (c == '\'')
				{
					if (isText && i + 1 < fullAddress.Length && fullAddress[i] == '\'')
					{
						if (hasSheet)
						{
							second += c;
						}
						else
						{
							first += c;
						}
					}

					isText = !isText;
				}
				else
				{
					if (brackPos.Count > 0)
					{
						if (c == '[' && !isText)
						{
							brackPos.Push(i);
						}
						else if (c == ']' && !isText)
						{
							if (brackPos.Count > 0)
							{
								var from = brackPos.Pop();
								bracketParts.Add(fullAddress.Substring(from + 1, i - from - 1));

								if (brackPos.Count == 0)
								{
									HandleBrackets(first, bracketParts);
								}
							}
							else
							{
								//Invalid address!
								return;
							}
						}
					}
					else if (c == '[' && !isText)
					{
						brackPos.Push(i);
					}
					else if (c == '!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
					{
						// the following is to handle addresses that specifies the
						// same worksheet twice: Sheet1!A1:Sheet1:A3
						// They will be converted to: Sheet1!A1:A3
						if (hasSheet && second != null && second.ToLower(CultureInfo.CurrentCulture).EndsWith(first.ToLower(CultureInfo.CurrentCulture)))
						{
							second = Regex.Replace(second, $"{first}$", string.Empty);
						}

						hasSheet = true;
					}
					else if (c == ',' && !isText)
					{
						isMulti = true;
						SetAddress(ref first, ref second, ref hasSheet, isMulti);
					}
					else
					{
						if (hasSheet)
						{
							second += c;
						}
						else
						{
							first += c;
						}
					}
				}
			}

			if (Table == null)
			{
				SetAddress(ref first, ref second, ref hasSheet, isMulti);
			}
		}
		catch
		{
			// An unparsable address leaves the range unset
		}
	}
	private void SetAddress(ref string first, ref string second, ref bool hasSheet, bool isMulti)
	{
		string ws, address;
		if (hasSheet)
		{
			ws = first;
			address = second;
			first = "";
			second = "";
		}
		else
		{
			address = first;
			ws = "";
			first = "";
		}

		hasSheet = false;
		if (string.IsNullOrEmpty(_firstAddress))
		{
			if (string.IsNullOrEmpty(_ws) || !string.IsNullOrEmpty(ws))
			{
				_ws = ws;
			}

			_firstAddress = address;
			GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
		}

		if (isMulti)
		{
			_addresses ??= [];
			_addresses.Add(new ExcelAddress(_ws, address));
		}
		else
		{
			_addresses = null;
		}
	}
}

using OfficeOpenXml.Utils;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml;

public sealed partial class ExcelWorkbook
{
	/// <summary>
	/// Read shared strings to list
	/// </summary>
	private void GetSharedStrings()
	{
		if (_package.Package.PartExists(SharedStringsUri))
		{
			var xml = _package.GetXmlFromUri(SharedStringsUri);
			var nl = xml.SelectNodes("//d:sst/d:si", NameSpaceManager);
			_sharedStringsList = [];
			if (nl != null)
			{
				foreach (XmlNode node in nl)
				{
					var n = node.SelectSingleNode("d:t", NameSpaceManager);
					if (n != null)
					{
						_sharedStringsList.Add(new SharedStringItem { Text = ConvertUtil.ExcelDecodeString(n.InnerText) });
					}
					else
					{
						_sharedStringsList.Add(new SharedStringItem { Text = node.InnerXml, isRichText = true });
					}
				}
			}
			//Delete the shared string part, it will be recreated when the package is saved.
			foreach (var rel in Part.GetRelationships())
			{
				if (rel.TargetUri.OriginalString.EndsWith("sharedstrings.xml", StringComparison.OrdinalIgnoreCase))
				{
					Part.DeleteRelationship(rel.Id);
					break;
				}
			}

			_package.Package.DeletePart(SharedStringsUri); //Remove the part, it is recreated when saved.
		}
	}
	internal void GetDefinedNames()
	{
		var nl = WorkbookXml.SelectNodes("//d:definedNames/d:definedName", NameSpaceManager);
		if (nl != null)
		{
			foreach (XmlElement elem in nl)
			{
				var fullAddress = elem.InnerText;

				ExcelWorksheet nameWorksheet;

				if (!int.TryParse(elem.GetAttribute("localSheetId"), NumberStyles.Number, CultureInfo.InvariantCulture, out var localSheetID))
				{
					localSheetID = -1;
					nameWorksheet = null;
				}
				else
				{
					nameWorksheet = Worksheets[localSheetID + _package._worksheetAdd];
				}

				var addressType = ExcelAddressBase.IsValid(fullAddress);
				ExcelNamedRange namedRange;

				fullAddress = ResolveExternalReferenceIndex(fullAddress);

				if (addressType is ExcelAddressBase.AddressType.Invalid or ExcelAddressBase.AddressType.InternalName or ExcelAddressBase.AddressType.ExternalName or ExcelAddressBase.AddressType.Formula or ExcelAddressBase.AddressType.ExternalAddress)    //A value or a formula
				{
					namedRange = AddValueOrFormulaName(elem, nameWorksheet, fullAddress);
				}
				else
				{
					namedRange = AddAddressName(elem, localSheetID, fullAddress);
				}

				if (elem.GetAttribute("hidden") == "1" && namedRange != null)
				{
					namedRange.IsNameHidden = true;
				}

				if (!string.IsNullOrEmpty(elem.GetAttribute("comment")))
				{
					namedRange.NameComment = elem.GetAttribute("comment");
				}
			}
		}
	}

	private string ResolveExternalReferenceIndex(string fullAddress)
	{
		if (fullAddress.StartsWith("[", StringComparison.CurrentCulture))
		{
			var start = fullAddress.IndexOf("[", StringComparison.CurrentCulture);
			var end = fullAddress.IndexOf("]", start, StringComparison.CurrentCulture);
			if (start >= 0 && end >= 0)
			{

				var externalIndex = fullAddress.Substring(start + 1, end - start - 1);
				if (int.TryParse(externalIndex, NumberStyles.Any, CultureInfo.InvariantCulture, out var index))
				{
					if (index > 0 && index <= _externalReferences.Count)
					{
						fullAddress = fullAddress[..start] + "[" + _externalReferences[index - 1] + "]" + fullAddress[(end + 1)..];
					}
				}
			}
		}

		return fullAddress;
	}

	private ExcelNamedRange AddValueOrFormulaName(XmlElement elem, ExcelWorksheet nameWorksheet, string fullAddress)
	{
		ExcelNamedRange namedRange;
		var range = new ExcelRangeBase(this, nameWorksheet, elem.GetAttribute("name"), true);
		if (nameWorksheet == null)
		{
			namedRange = _names.Add(elem.GetAttribute("name"), range);
		}
		else
		{
			namedRange = nameWorksheet.Names.Add(elem.GetAttribute("name"), range);
		}

		if (ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "\"")) //String value
		{
			namedRange.NameValue = fullAddress[1..^1];
		}
		else if (double.TryParse(fullAddress, NumberStyles.Number, CultureInfo.InvariantCulture, out var value))
		{
			namedRange.NameValue = value;
		}
		else
		{
			namedRange.NameFormula = fullAddress;
		}

		return namedRange;
	}

	private ExcelNamedRange AddAddressName(XmlElement elem, int localSheetID, string fullAddress)
	{
		ExcelAddress addr = new(fullAddress, _package, null);
		if (localSheetID > -1)
		{
			if (string.IsNullOrEmpty(addr._ws))
			{
				return Worksheets[localSheetID + _package._worksheetAdd].Names.Add(elem.GetAttribute("name"), new ExcelRangeBase(this, Worksheets[localSheetID + _package._worksheetAdd], fullAddress, false));
			}
			else
			{
				return Worksheets[localSheetID + _package._worksheetAdd].Names.Add(elem.GetAttribute("name"), new ExcelRangeBase(this, Worksheets[addr._ws], fullAddress, false));
			}
		}
		else
		{
			var ws = Worksheets[addr._ws];
			return _names.Add(elem.GetAttribute("name"), new ExcelRangeBase(this, ws, fullAddress, false));
		}
	}
}
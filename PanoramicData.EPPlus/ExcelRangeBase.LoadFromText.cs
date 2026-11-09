using OfficeOpenXml.Table;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Loads a CSV text into a range starting from the top left cell.
	/// Default settings is Comma separation
	/// </summary>
	/// <param name="Text">The Text</param>
	/// <returns>The range containing the data</returns>
	public ExcelRangeBase LoadFromText(string Text) => LoadFromText(Text, new ExcelTextFormat());
	/// <summary>
	/// Loads a CSV text into a range starting from the top left cell.
	/// </summary>
	/// <param name="Text">The Text</param>
	/// <param name="Format">Information how to load the text</param>
	/// <returns>The range containing the data</returns>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format)
	{
		if (string.IsNullOrEmpty(Text))
		{
			var r = _worksheet.Cells[_fromRow, _fromCol];
			r.Value = "";
			return r;
		}

		Format ??= new ExcelTextFormat();

		string[] lines;
		if (Format.TextQualifier == 0)
		{
			lines = Regex.Split(Text, Format.EOL);
		}
		else
		{
			lines = GetLines(Text, Format);
		}

		var row = 0;
		var col = 0;
		var maxCol = col;
		var lineNo = 1;
		var values = new List<object>[lines.Length];
		foreach (var line in lines)
		{
			var items = new List<object>();
			values[row] = items;

			if (lineNo > Format.SkipLinesBeginning && lineNo <= lines.Length - Format.SkipLinesEnd)
			{
				col = 0;
				var v = new StringBuilder();
				bool isText = false, isQualifier = false;
				var QCount = 0;
				var lineQCount = 0;
				foreach (var c in line)
				{
					if (Format.TextQualifier != 0 && c == Format.TextQualifier)
					{
						if (!isText && v.Length != 0)
						{
							throw new Exception(string.Format("Invalid Text Qualifier in line : {0}", line));
						}

						isQualifier = !isQualifier;
						QCount += 1;
						lineQCount++;
						isText = true;
					}
					else
					{
						if (QCount > 1 && v.Length != 0)
						{
							v.Append(Format.TextQualifier, QCount / 2);
						}
						else if (QCount > 2 && v.Length == 0)
						{
							v.Append(Format.TextQualifier, (QCount - 1) / 2);
						}

						if (isQualifier)
						{
							v.Append(c);
						}
						else
						{
							if (c == Format.Delimiter)
							{
								items.Add(ConvertData(Format, v.ToString(), col, isText));
								v.Clear();
								isText = false;
								col++;
							}
							else
							{
								if (QCount % 2 != 0)
								{
									throw new Exception(string.Format("Text delimiter is not closed in line : {0}", line));
								}

								v.Append(c);
							}
						}

						QCount = 0;
					}
				}

				if (QCount > 1 && v.Length != 0 && QCount == 2)
				{
					v.Append(Format.TextQualifier, QCount / 2);
				}

				if (lineQCount % 2 != 0)
				{
					throw new Exception(string.Format("Text delimiter is not closed in line : {0}", line));
				}

				items.Add(ConvertData(Format, v.ToString(), col, isText));
				if (col > maxCol)
				{
					maxCol = col;
				}

				row++;
			}

			lineNo++;
		}
		// flush
		_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + values.Length - 1, _fromCol + maxCol,
			(List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object _) =>
			{
				rowIx -= _fromRow;
				columnIx -= _fromCol;
				var item = values[rowIx];
				if (item == null || item.Count <= columnIx)
				{
					return;
				}

				list[index] = new ExcelCoreValue { _value = item[columnIx], _styleId = list[index]._styleId };
			}, values);

		return _worksheet.Cells[_fromRow, _fromCol, _fromRow + row - 1, _fromCol + maxCol];
	}

	/// <summary>
	/// Loads a CSV text into a range starting from the top left cell.
	/// </summary>
	/// <param name="Text">The Text</param>
	/// <param name="Format">Information how to load the text</param>
	/// <param name="TableStyle">Create a table with this style</param>
	/// <param name="FirstRowIsHeader">Use the first row as header</param>
	/// <returns></returns>
	public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
	{
		var r = LoadFromText(Text, Format);

		var tbl = _worksheet.Tables.Add(r, "");
		tbl.ShowHeader = FirstRowIsHeader;
		tbl.TableStyle = TableStyle;

		return r;
	}
	/// <summary>
	/// Loads a CSV file into a range starting from the top left cell.
	/// </summary>
	/// <param name="TextFile">The Textfile</param>
	/// <returns></returns>
	public ExcelRangeBase LoadFromText(FileInfo TextFile) => LoadFromText(File.ReadAllText(TextFile.FullName, Encoding.ASCII));
	/// <summary>
	/// Loads a CSV file into a range starting from the top left cell.
	/// </summary>
	/// <param name="TextFile">The Textfile</param>
	/// <param name="Format">Information how to load the text</param>
	/// <returns></returns>
	public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format) => LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format);
	/// <summary>
	/// Loads a CSV file into a range starting from the top left cell.
	/// </summary>
	/// <param name="TextFile">The Textfile</param>
	/// <param name="Format">Information how to load the text</param>
	/// <param name="TableStyle">Create a table with this style</param>
	/// <param name="FirstRowIsHeader">Use the first row as header</param>
	/// <returns></returns>
	public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader) => LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, TableStyle, FirstRowIsHeader);
}

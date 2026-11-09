using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System.Text;
using System.Xml;

namespace OfficeOpenXml;

public partial class ExcelRangeBase
{
	/// <summary>
	/// Set the hyperlink property for a range of cells
	/// </summary>
	public Uri Hyperlink
	{
		get
		{
			IsRangeValid("formulaR1C1");
			return _worksheet._hyperLinks.GetValue(_fromRow, _fromCol);
		}
		set
		{
			_changePropMethod(this, _setHyperLinkDelegate, value);
		}
	}
	/// <summary>
	/// If the value is in richtext format.
	/// </summary>
	public bool IsRichText
	{
		get
		{
			IsRangeValid("richtext");
			return _worksheet._flags.GetFlagValue(_fromRow, _fromCol, CellFlags.RichText);
		}
		set
		{
			_changePropMethod(this, _setIsRichTextDelegate, value);
		}
	}
	protected ExcelRichTextCollection? _rtc;
	/// <summary>
	/// Cell value is richtext formatted. 
	/// Richtext-property only apply to the left-top cell of the range.
	/// </summary>
	public ExcelRichTextCollection RichText
	{
		get
		{
			IsRangeValid("richtext");
			_rtc ??= GetRichText(_fromRow, _fromCol);

			return _rtc;
		}
	}

	private ExcelRichTextCollection GetRichText(int row, int col)
	{
		XmlDocument xml = new();
		var v = _worksheet.GetValueInner(row, col);
		var isRt = _worksheet._flags.GetFlagValue(row, col, CellFlags.RichText);
		if (v != null)
		{
			if (isRt)
			{
				XmlHelper.LoadXmlSafe(xml, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + v.ToString() + "</d:si>", Encoding.UTF8);
			}
			else
			{
				xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ><d:r><d:t>" + ConvertUtil.ExcelEscapeString(v.ToString()) + "</d:t></d:r></d:si>");
			}
		}
		else
		{
			xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
		}

		var rtc = new ExcelRichTextCollection(_worksheet.NameSpaceManager, xml.SelectSingleNode("d:si", _worksheet.NameSpaceManager), this);
		return rtc;
	}
	/// <summary>
	/// returns the comment object of the first cell in the range
	/// </summary>
	public ExcelComment Comment
	{
		get
		{
			IsRangeValid("comments");
			var i = -1;
			if (_worksheet.Comments.Count > 0)
			{
				if (_worksheet._commentsStore.Exists(_fromRow, _fromCol, ref i))
				{
					return _worksheet._comments[i] as ExcelComment;
				}
			}

			return null;
		}
	}
	/// <summary>
	/// Adds a new comment for the range.
	/// If this range contains more than one cell, the top left comment is returned by the method.
	/// </summary>
	/// <param name="Text"></param>
	/// <param name="Author"></param>
	/// <returns>A reference comment of the top left cell</returns>
	public ExcelComment AddComment(string Text, string Author)
	{
		if (string.IsNullOrEmpty(Author))
		{
			Author = System.Security.Claims.ClaimsPrincipal.Current.Identity.Name;
		}
		//Check if any comments exists in the range and throw an exception
		_changePropMethod(this, _setExistsCommentDelegate, null);
		//Create the comments
		_changePropMethod(this, _setCommentDelegate, new[] { Text, Author });

		return _worksheet.Comments[new ExcelCellAddress(_fromRow, _fromCol)];
	}
}

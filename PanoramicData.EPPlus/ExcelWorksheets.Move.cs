using System.Diagnostics.CodeAnalysis;

namespace OfficeOpenXml;

public partial class ExcelWorksheets
{
	/// <summary>
	/// Moves the source worksheet to the position before the target worksheet
	/// </summary>
	/// <param name="sourceName">The name of the source worksheet</param>
	/// <param name="targetName">The name of the target worksheet</param>
	public void MoveBefore(string sourceName, string targetName) => Move(sourceName, targetName, false);

	/// <summary>
	/// Moves the source worksheet to the position before the target worksheet
	/// </summary>
	/// <param name="sourcePositionId">The id of the source worksheet</param>
	/// <param name="targetPositionId">The id of the target worksheet</param>
	public void MoveBefore(int sourcePositionId, int targetPositionId) => Move(sourcePositionId, targetPositionId, false);

	/// <summary>
	/// Moves the source worksheet to the position after the target worksheet
	/// </summary>
	/// <param name="sourceName">The name of the source worksheet</param>
	/// <param name="targetName">The name of the target worksheet</param>
	public void MoveAfter(string sourceName, string targetName) => Move(sourceName, targetName, true);

	/// <summary>
	/// Moves the source worksheet to the position after the target worksheet
	/// </summary>
	/// <param name="sourcePositionId">The id of the source worksheet</param>
	/// <param name="targetPositionId">The id of the target worksheet</param>
	public void MoveAfter(int sourcePositionId, int targetPositionId) => Move(sourcePositionId, targetPositionId, true);

	/// <summary>
	/// 
	/// </summary>
	/// <param name="sourceName"></param>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public void MoveToStart(string sourceName)
	{
		var sourceSheet = this[sourceName];
		if (sourceSheet == null)
		{
			throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
		}

		Move(sourceSheet.PositionID, _pck._worksheetAdd, false);
	}

	/// <summary>
	/// 
	/// </summary>
	/// <param name="sourcePositionId"></param>
	public void MoveToStart(int sourcePositionId) => Move(sourcePositionId, _pck._worksheetAdd, false);

	/// <summary>
	/// 
	/// </summary>
	/// <param name="sourceName"></param>
	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	public void MoveToEnd(string sourceName)
	{
		var sourceSheet = this[sourceName];
		if (sourceSheet == null)
		{
			throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
		}

		Move(sourceSheet.PositionID, _worksheets.Count + (_pck._worksheetAdd - 1), true);
	}

	/// <summary>
	/// 
	/// </summary>
	/// <param name="sourcePositionId"></param>
	public void MoveToEnd(int sourcePositionId) => Move(sourcePositionId, _worksheets.Count + (_pck._worksheetAdd - 1), true);

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	private void Move(string sourceName, string targetName, bool placeAfter)
	{
		var sourceSheet = this[sourceName];
		if (sourceSheet == null)
		{
			throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
		}

		var targetSheet = this[targetName];
		if (targetSheet == null)
		{
			throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", targetName));
		}

		Move(sourceSheet.PositionID, targetSheet.PositionID, placeAfter);
	}

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	private void Move(int sourcePositionId, int targetPositionId, bool placeAfter)
	{
		// Bugfix: if source and target are the same worksheet the following code will create a duplicate
		//         which will cause a corrupt workbook. /swmal 2014-05-10
		if (sourcePositionId == targetPositionId)
		{
			return;
		}

		lock (_worksheets)
		{
			var sourceSheet = this[sourcePositionId];
			if (sourceSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", sourcePositionId));
			}

			var targetSheet = this[targetPositionId];
			if (targetSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", targetPositionId));
			}

			if (sourcePositionId == targetPositionId && _worksheets.Count < 2)
			{
				return;     //--- no reason to attempt to re-arrange a single item with itself
			}

			_worksheets = GetMovedOrder(sourcePositionId, targetPositionId, placeAfter, sourceSheet);

			MoveSheetXmlNode(sourceSheet, targetSheet, placeAfter);
		}
	}

	private Dictionary<int, ExcelWorksheet> GetMovedOrder(int sourcePositionId, int targetPositionId, bool placeAfter, ExcelWorksheet sourceSheet)
	{
		var index = _pck._worksheetAdd;
		var newOrder = new Dictionary<int, ExcelWorksheet>();
		foreach (var entry in _worksheets)
		{
			if (entry.Key == targetPositionId)
			{
				if (!placeAfter)
				{
					sourceSheet.PositionID = index;
					newOrder.Add(index++, sourceSheet);
				}

				entry.Value.PositionID = index;
				newOrder.Add(index++, entry.Value);

				if (placeAfter)
				{
					sourceSheet.PositionID = index;
					newOrder.Add(index++, sourceSheet);
				}
			}
			else if (entry.Key == sourcePositionId)
			{
				//--- do nothing
			}
			else
			{
				entry.Value.PositionID = index;
				newOrder.Add(index++, entry.Value);
			}
		}

		return newOrder;
	}

	[SuppressMessage("Major Code Smell", "S112:General or reserved exceptions should never be thrown", Justification = "Published behaviour: callers may catch this exact exception type, so it is kept.")]
	private void MoveSheetXmlNode(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, bool placeAfter)
	{
		lock (TopNode.OwnerDocument)
		{
			var sourceNode = TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", sourceSheet.SheetID), _namespaceManager);
			var targetNode = TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", targetSheet.SheetID), _namespaceManager);
			if (sourceNode == null || targetNode == null)
			{
				throw new Exception("Source SheetId and Target SheetId must be valid");
			}

			if (placeAfter)
			{
				TopNode.InsertAfter(sourceNode, targetNode);
			}
			else
			{
				TopNode.InsertBefore(sourceNode, targetNode);
			}
		}
	}
}

using OfficeOpenXml;

internal class FlagCellStore : CellStore<byte>
{
	internal void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
	{
		var currentValue = (CellFlags)GetValue(Row, Col);
		if (value)
		{
			SetValue(Row, Col, (byte)(currentValue | cellFlags)); // add the CellFlag bit
		}
		else
		{
			SetValue(Row, Col, (byte)(currentValue & ~cellFlags)); // remove the CellFlag bit
		}
	}
	internal bool GetFlagValue(int Row, int Col, CellFlags cellFlags) => ((byte)cellFlags & GetValue(Row, Col)) != 0;
}

using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing;

internal class DependencyChain
{
	internal List<FormulaCell> list = [];
	internal Dictionary<ulong, int> index = [];
	internal List<int> CalcOrder = [];
	internal void Add(FormulaCell f)
	{
		list.Add(f);
		f.Index = list.Count - 1;
		index.Add(ExcelCellBase.GetCellID(f.SheetID, f.Row, f.Column), f.Index);
	}
}
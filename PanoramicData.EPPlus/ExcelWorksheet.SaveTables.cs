using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.Text.RegularExpressions;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	/// <summary>
	/// Save all table data
	/// </summary>
	private void SaveTables()
	{
		foreach (var tbl in Tables)
		{
			if (tbl.ShowHeader || tbl.ShowTotal)
			{
				SaveTableColumns(tbl);
			}

			SaveTablePart(tbl);
		}
	}

	private void SaveTableColumns(ExcelTable tbl)
	{
		var colNum = tbl.Address._fromCol;
		var colVal = new HashSet<string>();
		foreach (var col in tbl.Columns)
		{
			var n = GetTableColumnName(tbl, col, colNum);

			if (colVal.Contains(n))
			{
				throw new InvalidDataException(string.Format("Table {0} Column {1} does not have a unique name.", tbl.Name, col.Name));
			}

			colVal.Add(n);
			if (!string.IsNullOrEmpty(col.CalculatedColumnFormula))
			{
				SetCalculatedColumnFormula(tbl, col, colNum);
			}

			colNum++;
		}
	}

	private string GetTableColumnName(ExcelTable tbl, ExcelTableColumn col, int colNum)
	{
		string n;
		if (tbl.ShowHeader)
		{
			n = tbl.WorkSheet.GetValue<string>(tbl.Address._fromRow,
				tbl.Address._fromCol + col.Position);
			if (string.IsNullOrEmpty(n))
			{
				n = col.Name.ToLowerInvariant();
				SetValueInner(tbl.Address._fromRow, colNum, ConvertUtil.ExcelDecodeString(col.Name));
			}
			else
			{
				col.Name = n;
			}
		}
		else
		{
			n = col.Name.ToLowerInvariant();
		}

		return n;
	}

	private void SetCalculatedColumnFormula(ExcelTable tbl, ExcelTableColumn col, int colNum)
	{
		var fromRow = tbl.ShowHeader ? tbl.Address._fromRow + 1 : tbl.Address._fromRow;
		var toRow = tbl.ShowTotal ? tbl.Address._toRow - 1 : tbl.Address._toRow;
		var r1c1Formula = ExcelCellBase.TranslateToR1C1(col.CalculatedColumnFormula, fromRow, colNum);
		var needsTranslation = r1c1Formula != col.CalculatedColumnFormula;

		for (var row = fromRow; row <= toRow; row++)
		{
			SetFormula(row, colNum, needsTranslation ? ExcelCellBase.TranslateFromR1C1(r1c1Formula, row, colNum) : r1c1Formula);
		}
	}

	private void SaveTablePart(ExcelTable tbl)
	{
		if (tbl.Part == null)
		{
			var id = tbl.Id;
			tbl.TableUri = GetNewUri(_package.Package, @"/xl/tables/table{0}.xml", ref id);
			tbl.Id = id;
			tbl.Part = _package.Package.CreatePart(tbl.TableUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", Workbook._package.Compression);
			var stream = tbl.Part.GetStream(FileMode.Create);
			tbl.TableXml.Save(stream);
			var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, tbl.TableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");
			tbl.RelationshipID = rel.Id;

			CreateNode("d:tableParts");
			var tbls = TopNode.SelectSingleNode("d:tableParts", NameSpaceManager);

			var tblNode = tbls.OwnerDocument.CreateElement("tablePart", ExcelPackage.schemaMain);
			tbls.AppendChild(tblNode);
			tblNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
		}
		else
		{
			var stream = tbl.Part.GetStream(FileMode.Create);
			tbl.TableXml.Save(stream);
		}
	}

	private static string GetTotalFunction(ExcelTableColumn col, string FunctionNum)
	{
		var escapedColumn = Regex.Replace(col.Name, @"[\[\]#']", new MatchEvaluator(m => "'" + m.Value));
		return string.Format("SUBTOTAL({0},{1}[{2}])", FunctionNum, col._tbl.Name, escapedColumn);
	}
}

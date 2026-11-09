using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Text;

namespace OfficeOpenXml;

public partial class ExcelWorksheet
{
	internal class Formulas
	{
		public Formulas(ISourceCodeTokenizer tokenizer)
		{
			_tokenizer = tokenizer;
		}

		private readonly ISourceCodeTokenizer _tokenizer;
		internal int Index { get; set; }
		internal string Address { get; set; }
		internal bool IsArray { get; set; }
		public string Formula { get; set; }
		public int StartRow { get; set; }
		public int StartCol { get; set; }

		private IEnumerable<Token> Tokens { get; set; }

		internal string GetFormula(int row, int column, string worksheet)
		{
			if (StartRow == row && StartCol == column)
			{
				return Formula;
			}

			Tokens ??= _tokenizer.Tokenize(Formula, worksheet);

			var f = new StringBuilder();
			foreach (var token in Tokens)
			{
				if (token.TokenType == TokenType.ExcelAddress)
				{
					var a = new ExcelFormulaAddress(token.Value);
					f.Append(!string.IsNullOrEmpty(a._wb) || !string.IsNullOrEmpty(a._ws)
						? token.Value
						: a.GetOffset(row - StartRow, column - StartCol));
				}
				else
				{
					f.Append(token.Value);
				}
			}

			return f.ToString();
		}

		public Formulas Clone()
		{
			var formulas = new Formulas(_tokenizer)
			{
				Index = Index,
				Address = Address,
				IsArray = IsArray,
				Formula = Formula,
				StartRow = StartRow,
				StartCol = StartCol
			};
			return formulas;
		}
	}
}

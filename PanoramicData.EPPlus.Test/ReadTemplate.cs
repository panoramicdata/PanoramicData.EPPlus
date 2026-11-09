using OfficeOpenXml;
using OfficeOpenXml.Style;
namespace PanoramicData.EPPlus.Test;

[TestClass]
public partial class ReadTemplate
{
	[TestMethod]
	public void ReadBlankStream()
	{
		MemoryStream stream = new();
		using (ExcelPackage pck = new(stream))
		{
			pck.Workbook.Worksheets.Add("Perf");
			pck.SaveAs(stream);
		}

		stream.Close();
	}
	public static byte[] ReadTemplateFile(string templateName)
	{
		byte[] templateFIle;
		using (MemoryStream ms = new())
		{
			using (var sw = new FileStream(templateName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				var buffer = new byte[2048];
				int bytesRead;
				while ((bytesRead = sw.Read(buffer, 0, buffer.Length)) > 0)
				{
					ms.Write(buffer, 0, bytesRead);
				}
			}

			ms.Position = 0;
			templateFIle = ms.ToArray();
		}

		return templateFIle;
	}

	private static void CreateXlsxSheet(string pFileName, int pRows, int pColumns)
	{
		if (File.Exists(pFileName))
		{
			File.Delete(pFileName);
		}

		using ExcelPackage excelPackage = new(new FileInfo(pFileName));
		var excelWorksheet = excelPackage.Workbook.Worksheets.Add("Testsheet");

		// Fill with data
		for (var row = 1; row <= pRows; row++)
		{
			for (var column = 1; column <= pColumns; column++)
			{
				if (column > 1 && row > 2)
				{
					using (var range = excelWorksheet.Cells[row, column])
					{
						range.Style.Numberformat.Format = "0";
						range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
						range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					}

					excelWorksheet.Cells[row, column].Value = row * column;
				}
			}
		}

		// Try to style the first column, begining with row 3 which has no content yet...
		using (var range = excelWorksheet.Cells[ExcelCellBase.GetAddress(3, 1, pRows, 1)])
		{
			_ = range.Style;
		}

		// now I would add data to the first column (left out here)...
		excelPackage.Save();
	}
}

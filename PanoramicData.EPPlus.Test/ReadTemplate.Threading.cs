using OfficeOpenXml;
namespace PanoramicData.EPPlus.Test;

public partial class ReadTemplate
{
	#region "Threading Cellstore Test"
	private int _threadCount;
	[TestMethod, Ignore]
	public void ThreadingTest()
	{
		var pckThread = new ExcelPackage();
		var ws = pckThread.Workbook.Worksheets.Add("Threading");

		for (var t = 0; t < 20; t++)
		{
			var ts = new ThreadState(Finnished)
			{
				ws = ws,
				StartRow = 1 + t * 1000,
				Rows = 1000,
			};
			var tstart = new ThreadStart(ts.StartLoad);
			var thread = new Thread(tstart);
			_threadCount++;
			thread.Start();
		}

		while (true)
		{
			if (_threadCount == 0)
			{
				pckThread.SaveAs(new FileInfo("c:\\temp\\thread.xlsx"));
				break;
			}

			Thread.Sleep(1000);
		}
	}
	public void Finnished() => _threadCount--;
	private class ThreadState
	{
		public ThreadState(cbFinished cb)
		{
			_cb = cb;
		}
		public ExcelWorksheet ws { get; set; }
		public int StartRow { get; set; }
		public int Rows { get; set; }
		public delegate void cbFinished();
		public cbFinished _cb;
		public void StartLoad()
		{
			for (var row = StartRow; row < StartRow + Rows; row++)
			{
				for (var col = 1; col < 100; col++)
				{
					ws.SetValue(row, col, string.Format("row {0} col {1}", row, col));
				}
			}

			_cb();
		}
	}
	#endregion
	[Ignore]
	[TestMethod]
	public void TestInvalidVBA()
	{
		const string infile = @"C:\temp\bug\Infile.xlsm";
		const string outfile = @"C:\temp\bug\Outfile.xlsm";
		ExcelPackage ep;

		using (var fs = File.OpenRead(infile))
		{
			ep = new ExcelPackage(fs);
		}

		using (var fs = File.OpenWrite(outfile))
		{
			ep.SaveAs(fs);
		}

		using (var fs = File.OpenRead(outfile))
		{
			ep = new ExcelPackage(fs);
		}

		using (var fs = File.OpenWrite(outfile))
		{
			ep.SaveAs(fs);
		}
	}
}

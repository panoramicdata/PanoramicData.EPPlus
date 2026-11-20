
namespace PanoramicData.EPPlus.Test;

/// <summary>
/// Summary description for UnitTest1
/// </summary>
[TestClass]
public partial class DrawingTest : TestBase
{
	[TestMethod]
	public void RunDrawingTests()
	{
		BarChart();
		Column();
		Cone();
		Dougnut();
		Drawings();
		Line();
		LineMarker();
		PieChart();
		PieChart3D();
		PieOfChart();
		Pyramid();
		Scatter();
		Bubble();
		Radar();
		Surface();
		Line2Test();
		MultiChartSeries();
		Picture();
		DrawingRowheightDynamic();
		DrawingSizingAndPositioning();
		DeleteDrawing();

		SaveWorksheet("Drawing.xlsx");

		ReadDocument();
		ReadDrawing();
	}
}


namespace PanoramicData.EPPlus.Test;

/// <summary>
/// This class contains testcases for issues on Codeplex and Github.
/// All tests requiering an template should be set to ignored as it's not practical to include all xlsx templates in the project.
/// </summary>
[TestClass]
public partial class Issues : TestBase
{
	[TestInitialize]
	public void Initialize()
	{
		if (!Directory.Exists(@"c:\Temp"))
		{
			Directory.CreateDirectory(@"c:\Temp");
		}

		if (!Directory.Exists(@"c:\Temp\bug"))
		{
			Directory.CreateDirectory(@"c:\Temp\bug");
		}
	}

	public class cls1
	{
		public int prop1 { get; set; }
	}

	public class cls2 : cls1
	{
		public string prop2 { get; set; }
	}
	public class Car
	{
		public int Id { get; set; }
		public string Make { get; set; }
		public string Model { get; set; }
		public int ReleaseYear { get; set; }

		public Car(int id, string make, string model, int releaseYear)
		{
			Id = id;
			Make = make;
			Model = model;
			ReleaseYear = releaseYear;
		}

		internal static List<Car> GenerateList() => [
			//random data
			new(1,"Toyota", "Carolla", 1950),
			new(2,"Toyota", "Yaris", 2000),
			new(3,"Toyota", "Hilux", 1990),
			new(4,"Nissan", "Juke", 2010),
			new(5,"Nissan", "Trail Blazer", 1995),
			new(6,"Nissan", "Micra", 2018),
			new(7,"BMW", "M3", 1980),
			new(8,"BMW", "X5", 2008),
			new(9,"BMW", "M6", 2003),
			new(10,"Merc", "S Class", 2001)
		];
	}
}
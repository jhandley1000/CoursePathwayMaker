using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("File Path For Data File:");
			var dataFilePath = Console.ReadLine();

			Console.WriteLine("Start Year:");
			var startYear = Convert.ToInt32(Console.ReadLine());

			Console.WriteLine("End Year:");
			var endYear = Convert.ToInt32(Console.ReadLine());

			Console.WriteLine("Campus:");
			var campus = Console.ReadLine();

			var excelHandler = new ExcelHandler(dataFilePath, campus);
			excelHandler.SetUpPathwayFile(startYear, endYear);
			excelHandler.BuildPathwaysForEachStudent(endYear-startYear+1);
			excelHandler.SavePathwayFile();
			excelHandler.QuitApp();
		}
	}
}

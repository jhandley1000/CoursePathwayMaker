using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker
{
	class ExcelHandler
	{
		public Application App { get; }
		public Workbook DataWorkbook { get; }
		public Workbook PathwayWorkbook { get; set; }
		List<int> StudentIDs { get; set; }
		string campus { get; }

		public ExcelHandler(string dataFilePath, string campus)
		{
			App = new Application();
			DataWorkbook = App.Workbooks.Open(dataFilePath);
			this.campus = campus;
			GetStudentIDs();
		}

		public void SetUpPathwayFile(int startYear, int endYear)
		{
			try
			{
				PathwayWorkbook = App.Workbooks.Add();
				var worksheet = PathwayWorkbook.Worksheets.Add();
				worksheet.Name = campus;

				worksheet.Cells[2, 1] = "Students";
				var yearCount = startYear;
				var colCount = 2;
				while (yearCount <= endYear)
				{
					worksheet.Cells[1, colCount] = yearCount;
					worksheet.Cells[1, colCount + 1] = yearCount;
					worksheet.Cells[2, colCount] = 1;
					worksheet.Cells[2, colCount + 1] = 2;
					yearCount += 1;
					colCount += 2;
				}

				PathwayWorkbook.SaveAs(@"E:\Dev\xlStuff\PathwayFile.xlsx");
			}
			catch
			{
				Console.WriteLine("Failed setting up Pathway File.");
				App.Quit();
			}
		}

		void GetStudentIDs()
		{
			StudentIDs = new List<int>();

			try
			{
				var studentWorksheet = DataWorkbook.Worksheets.get_Item("Students");

				var count = 1;
				foreach (var row in studentWorksheet.UsedRange.Rows)
				{
					StudentIDs.Add(Convert.ToInt32(studentWorksheet.Cells[count, 1].Value));
					count += 1;
				}
			}
			catch
			{
				App.Quit();
			}
		}

		public void SavePathwayFile()
		{
			try
			{
				PathwayWorkbook.Save();
			}
			catch
			{
				Console.WriteLine("Failed to Save Pathway File.");
				App.Quit();
			}
		}

		public void BuildPathwaysForEachStudent(int numYears)
		{
			try
			{
				var pathwayWorksheet = PathwayWorkbook.Worksheets.get_Item(campus);
				var dataWorksheet = DataWorkbook.Worksheets.get_Item("EnrollmentData");

				var pathwayRowCount = 3;
				foreach (var student in StudentIDs)
				{
					Console.WriteLine("{0}", student);
					var count = 0;
					while (count < numYears*2)
					{
						Console.WriteLine("---- {0} ----", count);
						var rowNum = 1;
						foreach (var row in dataWorksheet.UsedRange.Rows)
						{
							Console.WriteLine("{0}", rowNum);
							if ((Convert.ToInt32(dataWorksheet.Cells[rowNum + 1, 1].Value).Equals(Convert.ToInt32(pathwayWorksheet.Cells[1, count + 2].Value)))
								&& (Convert.ToInt32(dataWorksheet.Cells[rowNum + 1, 2].Value).Equals(Convert.ToInt32(pathwayWorksheet.Cells[2, count + 2].Value)))
								&& (Convert.ToInt32(dataWorksheet.Cells[rowNum + 1, 4].Value).Equals(Convert.ToInt32(student))))
							{ 
								pathwayWorksheet.Cells[pathwayRowCount, count + 1].Value += dataWorksheet.Cells[rowNum, 3];
							}
							rowNum += 1;
						}
						count += 1;
					}
					pathwayRowCount += 1;
				}
			}
			catch (Exception ex)
			{
				App.Quit();
			}
			
		}

		public void QuitApp()
		{
			App.Quit();
		}
	}
}

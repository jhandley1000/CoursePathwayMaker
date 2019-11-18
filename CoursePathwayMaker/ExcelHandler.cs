using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker
{
    public class ExcelHandler
    {
        public Application App { get; }
        public Workbook DataWorkbook { get; }
        public Workbook PathwayWorkbook { get; set; }
        public Workbook OutputWorkbook { get; set; }
        List<Workbook> Workbooks { get; set; }

        string pathwayFilePath { get; }
        List<Student> Students { get; }
		string campus { get; }

		public ExcelHandler(string dataFilePath, string campus, string pathwayFilePath)
		{
			App = new Application();
			DataWorkbook = App.Workbooks.Open(dataFilePath);
			this.campus = campus;
            Students = new List<Student>();
            this.pathwayFilePath = pathwayFilePath;
			//GetStudentIDs();
		}
        
        public ExcelHandler(List<string> filePaths)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            
            App = new Application();
            App.Visible = true;
            App.UserControl = true;

            Workbooks = new List<Workbook>();
            //PathwayWorkbook = App.Workbooks.Open(pathwayFilePath);
            //OutputWorkbook = App.Workbooks.Open(outPutFilePath);
            foreach (var filePath in filePaths)
            {
                Workbooks.Add(App.Workbooks.Open(filePath));
            }
        }

        public Worksheet GetWorksheet(int index, string sheetName)
        {
            return (Workbooks[index].Worksheets.get_Item(sheetName) as Worksheet);
        }

		public void SetUpPathwayFile(int startYear, int endYear)
		{
			try
			{
				PathwayWorkbook = App.Workbooks.Add();
				var worksheet = PathwayWorkbook.Worksheets.Add();
				(worksheet as Worksheet).Name = campus;

				(worksheet as Worksheet).Cells[2, 1] = "Students";
				var yearCount = startYear;
				var colCount = 2;
				while (yearCount <= endYear)
				{
					(worksheet as Worksheet).Cells[1, colCount] = yearCount;
					(worksheet as Worksheet).Cells[1, colCount + 1] = yearCount;
					(worksheet as Worksheet).Cells[2, colCount] = 1;
					(worksheet as Worksheet).Cells[2, colCount + 1] = 2;
					yearCount += 1;
					colCount += 2;
				}

				PathwayWorkbook.SaveAs(pathwayFilePath);
			}
			catch
			{
				Console.WriteLine("Failed setting up Pathway File.");
				App.Quit();
                throw;
			}
		}

		List<int> GetStudentIDs()
		{
			var studentIDs = new List<int>();

			try
			{
				var studentWorksheet = DataWorkbook.Worksheets.get_Item("Students");

				var count = 1;
				foreach (var row in (studentWorksheet as Worksheet).UsedRange.Rows)
				{
					studentIDs.Add(Convert.ToInt32(((studentWorksheet as Worksheet).Cells[count, 1] as Range).Value));
					count += 1;
				}

                return studentIDs;
			}
			catch
			{
                Console.WriteLine("Error: Problem Accessing Student IDs.");
				App.Quit();
                throw;
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
                throw;
			}
		}

		//public void BuildPathwaysForEachStudent(int numYears)
		//{
		//	try
		//	{
		//		var pathwayWorksheet = PathwayWorkbook.Worksheets.get_Item(campus);
		//		var dataWorksheet = DataWorkbook.Worksheets.get_Item("EnrollmentData");

		//		var pathwayRowCount = 3;
		//		foreach (var student in StudentIDs)
		//		{
		//			Console.WriteLine("{0}", student);
  //                  pathwayWorksheet.Cells[pathwayRowCount, 1] = student;
		//			var count = 0;
		//			while (count < numYears*2)
		//			{
		//				Console.WriteLine("---- {0} S{1}", pathwayWorksheet.Cells[1, count + 2].Value, pathwayWorksheet.Cells[2, count + 2].Value);
		//				var rowNum = 1;
		//				foreach (var row in dataWorksheet.UsedRange.Rows)
		//				{
		//					if ((Convert.ToInt32(dataWorksheet.Cells[rowNum + 1, 1].Value).Equals(Convert.ToInt32(pathwayWorksheet.Cells[1, count + 2].Value)))
		//						&& (Convert.ToInt32(dataWorksheet.Cells[rowNum + 1, 2].Value).Equals(Convert.ToInt32(pathwayWorksheet.Cells[2, count + 2].Value)))
		//						&& (Convert.ToInt32(dataWorksheet.Cells[rowNum + 1, 4].Value).Equals(Convert.ToInt32(student))))
		//					{ 
  //                              if (pathwayWorksheet.Cells[pathwayRowCount, count + 2].Value != null && pathwayWorksheet.Cells[pathwayRowCount, count + 2].Value.StartsWith("System.__ComObject"))
  //                              {
  //                                  pathwayWorksheet.Cells[pathwayRowCount, count + 2].Value = "";
  //                              }
		//						pathwayWorksheet.Cells[pathwayRowCount, count + 2].Value += dataWorksheet.Cells[rowNum + 1, 3].Value + " ";
		//					}
		//					rowNum += 1;
		//				}
		//				count += 1;
		//			}
		//			pathwayRowCount += 1;
		//		}
		//	}
		//	catch (Exception ex)
		//	{
		//		App.Quit();
		//	}
			
		//}

        public void SearchSpreadsheetForStudentEnrollments()
        {
            try
            {
                Console.WriteLine("Searching enrollment data for each student...");
                var dataWorksheet = DataWorkbook.Worksheets.get_Item("EnrollmentData");
                var studentIDs = GetStudentIDs();
                foreach (var student in studentIDs)
                {
                    Console.WriteLine("{0}", student);
                    var newStudent = new Student(student);

                    var rowNum = 2;
                    foreach (var row in (dataWorksheet as Worksheet).UsedRange.Rows)
                    {
                        if (Convert.ToInt32(((dataWorksheet as Worksheet).Cells[rowNum, 4] as Range).Value).Equals(newStudent.StudentID))
                        {
                            newStudent.AddSubjectEnrollment(
                                new SubjectEnrollment(
                                    ((dataWorksheet as Worksheet).Cells[rowNum, 3] as Range).Value + " ",
                                    Convert.ToInt32(((dataWorksheet as Worksheet).Cells[rowNum, 1] as Range).Value),
                                    Convert.ToInt32(((dataWorksheet as Worksheet).Cells[rowNum, 2] as Range).Value)));
                        }
                        rowNum++;
                    }

                    Students.Add(newStudent);
                }
            }
            catch (Exception ex)
            {
                App.Quit();
                throw;
            }
        }

        public void GeneratePathwaysFromStudentData(int numYears)
        {
            try
            {
                Console.WriteLine("Generating pathways...");
                var pathwayWorksheet = PathwayWorkbook.Worksheets.get_Item(campus);

                var pathwayRowCount = 3;
                foreach (var student in Students)
                {
                    (pathwayWorksheet as Worksheet).Cells[pathwayRowCount, 1] = student.StudentID;
                    var count = 0;
                    while (count < numYears*2)
                    {
                        foreach (var enrollment in student.SubjectEnrollments)
                        {
                            if (enrollment.Year.Equals(Convert.ToInt32(((pathwayWorksheet as Worksheet).Cells[1, count + 2] as Range).Value))
                                && (enrollment.Semester.Equals(Convert.ToInt32(((pathwayWorksheet as Worksheet).Cells[2, count + 2] as Range).Value))))
                            {
                                if (((pathwayWorksheet as Worksheet).Cells[pathwayRowCount, count + 2] as Range).Value != null && (((pathwayWorksheet as Worksheet).Cells[pathwayRowCount, count + 2] as Range).Value as string).StartsWith("System.__ComObject"))
                                {
                                    ((pathwayWorksheet as Worksheet).Cells[pathwayRowCount, count + 2] as Range).Value = "";
                                }
                                ((pathwayWorksheet as Worksheet).Cells[pathwayRowCount, count + 2] as Range).Value += enrollment.CourseCode;
                            }
                        }
                        count++;
                    }
                    pathwayRowCount++;
                }
            }
            catch
            {
                App.Quit();
                throw;
            }
            
        }

		public void QuitApp()
		{
			App.Quit();
		}
	}
}

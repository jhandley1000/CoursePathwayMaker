using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CoursePathwayMaker.ForceDirectedTableMaker
{
    public class ForceDirectedTableExcelHandler
    {
        string pathwayFilePath { get; }
        string forceDirectedTableFilePath { get; }
        Application App { get; }
        Workbook pathwayWorkbook { get; }
        Workbook forceDirectedTableWorkbook { get; set; }
        List<Connection> connections { get; set; }

        public ForceDirectedTableExcelHandler(string pathwayFilePath, string forceDirectedTableFilePath)
        {
            this.pathwayFilePath = pathwayFilePath;
            this.forceDirectedTableFilePath = forceDirectedTableFilePath;
            App = new Application();
            connections = new List<Connection>();
            pathwayWorkbook = App.Workbooks.Open(pathwayFilePath);
            forceDirectedTableWorkbook = SetUpForceDirectedTableWorkbook(forceDirectedTableFilePath);
        }

        Workbook SetUpForceDirectedTableWorkbook(string filePath)
        {
            try
            {
                var workbook = App.Workbooks.Add();
                var worksheet = workbook.Worksheets.Add();
                (worksheet as Worksheet).Name = "ForceDirectedTable";

                (worksheet as Worksheet).Cells[1, 1] = "From Course";
                (worksheet as Worksheet).Cells[1, 2] = "To Course";
                (worksheet as Worksheet).Cells[1, 3] = "Pathway Frequency";

                workbook.SaveAs(filePath);
                return workbook;
            }
            catch
            {
                App.Quit();
                throw;
            }
        }

        public void GetConnectionsFromPathwayFile()
        {
            try
            {
                var pathwayWorksheet = pathwayWorkbook.Worksheets.get_Item("Pathways");
                var pathwayRowNum = 3;
                foreach (var row in (pathwayWorksheet as Worksheet).UsedRange.Rows)
                {
                    var pathwayColNum = 0;
                    foreach (var cell in (pathwayWorksheet as Worksheet).UsedRange.Columns)
                    {
                        if (pathwayColNum >= 2)
                        {
                            var fromCourseCellValue = ((pathwayWorksheet as Worksheet).Cells[pathwayRowNum, pathwayColNum] as Range).Value;
                            var toCourseCellValue = ((pathwayWorksheet as Worksheet).Cells[pathwayRowNum, pathwayColNum + 1] as Range).Value;
                            if (fromCourseCellValue != null && toCourseCellValue != null)
                            {
								var fromCourseCodes = new List<string>((fromCourseCellValue as string).Trim().Split());
                                var toCourseCodes = new List<string>((toCourseCellValue as string).Trim().Split());

                                AddConnections(fromCourseCodes, toCourseCodes);
                            }
                        }
                        pathwayColNum++;
                    }
					pathwayRowNum++;
                }
            }
            catch
            {
                App.Quit();
                throw;
            }
        }

        void AddConnections(List<string> fromCourses, List<string> toCourses)
        {
            foreach (var fromCourseCode in fromCourses)
            {
                foreach (var toCourseCode in toCourses)
                {
                    AddConnectionOrIncreaseCount(fromCourseCode, toCourseCode);
                }
            }
        }

        void AddConnectionOrIncreaseCount(string fromCourseCode, string toCourseCode)
        {
            var connectionFound = false;
            foreach (var existingConnection in connections)
            {
                if (fromCourseCode.Equals(existingConnection.FromCourseCode) && toCourseCode.Equals(existingConnection.ToCourseCode))
                {
                    existingConnection.AddToPathwayFrequency();
                    connectionFound = true;
                }
            }

            if (!connectionFound)
            {
                connections.Add(new Connection(fromCourseCode, toCourseCode));
            }
        }

		public void PutAllConnectionsInOutputFile()
		{
			try
			{
				var worksheet = forceDirectedTableWorkbook.Worksheets.get_Item("ForceDirectedTable");
				var rowNum = 2;
				foreach (var connection in connections)
				{
					(worksheet as Worksheet).Cells[rowNum, 1] = connection.FromCourseCode;
					(worksheet as Worksheet).Cells[rowNum, 2] = connection.ToCourseCode;
					(worksheet as Worksheet).Cells[rowNum, 3] = connection.PathwayFrequency;

					rowNum++;
				}
			}
			catch
			{
				App.Quit();
				throw;
			}
		}

		public void SaveForceDirectedWorkbook()
		{
			try
			{
				forceDirectedTableWorkbook.Save();
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

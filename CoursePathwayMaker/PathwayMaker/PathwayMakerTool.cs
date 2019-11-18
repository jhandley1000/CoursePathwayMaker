using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CoursePathwayMaker.PathwayMaker
{
    public class PathwayMakerTool
    {
        public PathwayMakerTool()
        {
        }

        public void MakePathways(IConsoleReader consoleReader)
        {
            var dataFilePath = consoleReader.GetDataFilePath();
            var startYear = consoleReader.GetStartYear();
            var endYear = consoleReader.GetEndYear();
            var campus = consoleReader.GetCampus();
            var pathwayFilePath = consoleReader.GetFileSavePath();

            var excelHandler = new ExcelHandler(dataFilePath, campus, pathwayFilePath);

            excelHandler.SetUpPathwayFile(startYear, endYear);
            excelHandler.SearchSpreadsheetForStudentEnrollments();
            excelHandler.GeneratePathwaysFromStudentData(endYear - startYear + 1);

            Console.WriteLine("Saving {0}", pathwayFilePath);
            excelHandler.SavePathwayFile();
            excelHandler.QuitApp();
            Console.WriteLine("COMPLETE");
        }

        public List<Worksheet>GrabThisIsAWrapper(List<string> filePath)
        {
            var excelHandler = new ThisIsAWrapper(filePath);
            var worksheets = new List<Worksheet>();
            worksheets.Add(excelHandler.DoTheThing(0, "Ourimbah"));
            worksheets.Add(excelHandler.DoTheThing(1, "Ourimbah"));
            excelHandler.QuitApp();
            return worksheets;
        }
    }
}

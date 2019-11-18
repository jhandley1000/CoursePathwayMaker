using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CoursePathwayMaker.PathwayMaker
{
    public class ExcelFileManagerForTests
    {
        Application app;
        List<Workbooks> workbooks;

        public ExcelFileManagerForTests()
        {
            app = new Application() { Visible = true };
        }

        public Worksheet GetExcelWorksheet(string filePath, string worksheetName)
        {
            var workbook = app.Workbooks.Open(filePath);
            Worksheet worksheet = workbook.Worksheets[worksheetName] as Worksheet;
            return workbook.Worksheets[worksheetName] as Worksheet;
        }

        public void DeleteWorkbook(string filePath)
        {
            File.Delete(filePath);
        }

        public void QuitApp()
        {
            app.Quit();
        }
    }
}

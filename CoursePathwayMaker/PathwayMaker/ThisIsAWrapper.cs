using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.PathwayMaker
{
    public class ThisIsAWrapper
    {
        ExcelHandler handler;
        public ThisIsAWrapper(List<string> filePaths)
        {
            handler = new ExcelHandler(filePaths);
        }

        public Microsoft.Office.Interop.Excel.Worksheet DoTheThing(int index, string campus)
        {
            return handler.GetWorksheet(index, campus);
        }

        public void QuitApp()
        {
            handler.QuitApp();
        }
    }
}

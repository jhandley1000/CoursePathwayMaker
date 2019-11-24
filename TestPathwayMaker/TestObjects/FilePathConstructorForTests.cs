using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestPathwayMaker.TestObjects
{
    public class FilePathConstructorForTests
    {
        public string ConstructExcelFilePath(string folder, string filename)
        {
            return string.Format("{0}..\\..\\{1}\\{2}.xlsx", AppDomain.CurrentDomain.BaseDirectory, folder, filename);
        }
    }
}

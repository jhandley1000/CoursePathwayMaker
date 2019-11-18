using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoursePathwayMaker.PathwayMaker;

namespace UnitTestProject1.TestObjects
{
    public class ConsoleReaderForTests : IConsoleReader
    {
        string dataFilePath { get; }
        int startYear { get; }
        int endYear { get; }
        string campus { get; }
        string fileSavePath { get; }

        public ConsoleReaderForTests(string dataFilePath, int startYear, int endYear, string campus, string fileSavePath)
        {
            this.dataFilePath = dataFilePath;
            this.startYear = startYear;
            this.endYear = endYear;
            this.campus = campus;
            this.fileSavePath = fileSavePath;
        }

        public string GetDataFilePath()
        {
            return dataFilePath;
        }

        public int GetStartYear()
        {
            return startYear;
        }

        public int GetEndYear()
        {
            return endYear;
        }

        public string GetCampus()
        {
            return campus;
        }

        public string GetFileSavePath()
        {
            return fileSavePath;
        }
    }
}

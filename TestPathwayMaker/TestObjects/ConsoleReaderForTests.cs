using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoursePathwayMaker.PathwayMaker;

namespace TestPathwayMaker.TestObjects
{
    public class ConsoleReaderForTests : IConsoleReader
    {
        string inputFilePath { get; }
        int startYear { get; }
        int endYear { get; }
        string campus { get; }
        string fileSavePath { get; }
        string testFilePath { get; }
        string subjectArea { get; }
        string term { get; }
        string semester { get; }

        public ConsoleReaderForTests(string dataFilePath, int startYear, int endYear, string campus, string testFilePath)
        {
            this.startYear = startYear;
            this.endYear = endYear;
            this.campus = campus;
            this.fileSavePath = constructFileSavePath(testFilePath);
            this.testFilePath = constructTestFilePath(testFilePath);
            this.inputFilePath = constructDataFilePath(dataFilePath);
        }

        public ConsoleReaderForTests(string inputFilename, string saveFilePath, string testFilename)
        {
            this.inputFilePath = new FilePathConstructorForTests().ConstructExcelFilePath("PathwayFiles", inputFilename);
            this.fileSavePath = new FilePathConstructorForTests().ConstructExcelFilePath("OUTPUTFORTEST", saveFilePath);
            this.testFilePath = new FilePathConstructorForTests().ConstructExcelFilePath("TestTables", testFilename);
        }

        public ConsoleReaderForTests(string subjectArea, string semester, int year, string term)
        {
            this.subjectArea = subjectArea;
            this.term = term;
            this.semester = semester;
            this.startYear = year;
        }

        public string GetTerm()
        {
            return term;
        }

        public string GetSemester()
        {
            return semester;
        }

        public string GetSubjectArea()
        {
            return subjectArea;
        }

        public int GetStartYear()
        {
            return startYear;
        }

        public int GetEndYear()
        {
            return endYear;
        }

        public string GetWorksheetName()
        {
            return campus;
        }

        public string GetInputFilePath()
        {
            return inputFilePath;
        }

        public string GetNewSaveFilePath()
        {
            return fileSavePath;
        }

        public string GetTestFilePath()
        {
            return testFilePath;
        }

        string constructDataFilePath(string filename)
        {
            return new FilePathConstructorForTests().ConstructExcelFilePath("TestDataFile", filename);
        }

        string constructFileSavePath(string filename)
        {
            return new FilePathConstructorForTests().ConstructExcelFilePath("TestExcelFilesOUTPUT", filename + "OUTPUT");
        }

        string constructTestFilePath(string filename)
        {
            return new FilePathConstructorForTests().ConstructExcelFilePath("TestExcelFiles", filename);
        }
        public bool AddToDb()
        {
            return false;
        }
    }
}

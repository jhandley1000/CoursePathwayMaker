using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.PathwayMaker
{
    public class ConsoleReader : IConsoleReader
    {
        public string GetInputFilePath()
        {
            return GetInputFromUser("Input File Path: ");
        }

        public int GetStartYear()
        {
            return Convert.ToInt32(GetInputFromUser("Start Year: "));
        }

        public int GetEndYear()
        {
            return Convert.ToInt32(GetInputFromUser("End Year: "));
        }

        public string GetWorksheetName()
        {
            return GetInputFromUser("Worksheet Name: ");
        }

        public string GetNewSaveFilePath()
        {
            return GetInputFromUser("Save New File As: ");
        }

        public string GetTerm()
        {
            return GetInputFromUser("Term: ");
        }

        public string GetSemester()
        {
            return GetInputFromUser("Semester: ");
        }

        public string GetSubjectArea()
        {
            return GetInputFromUser("Subject Area: ");
        }

        public bool AddToDb()
        {
            var addToDb = false;

            if (GetInputFromUser("Add To Database (y/n): ").Equals("y"))
            {
                addToDb = true;
            }

            return addToDb;
        }

        string GetInputFromUser(string prompt)
        {
            Console.Write(prompt);
            return Console.ReadLine();
        }
    }
}

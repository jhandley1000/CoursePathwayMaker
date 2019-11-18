using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.PathwayMaker
{
    public class ConsoleReader : IConsoleReader
    {
        public string GetDataFilePath()
        {
            return GetInputFromUser("File Path For Data File: ");
        }

        public int GetStartYear()
        {
            return Convert.ToInt32(GetInputFromUser("Start Year: "));
        }

        public int GetEndYear()
        {
            return Convert.ToInt32(GetInputFromUser("End Year: "));
        }

        public string GetCampus()
        {
            return GetInputFromUser("Campus: ");
        }

        public string GetFileSavePath()
        {
            return GetInputFromUser("Save Pathway File As: ");
        }

        string GetInputFromUser(string prompt)
        {
            Console.Write(prompt);
            return Console.ReadLine();
        }
    }
}

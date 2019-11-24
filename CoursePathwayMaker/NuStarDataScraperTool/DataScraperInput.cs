using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.NuStarDataScraperTool
{
    public class DataScraperInput
    {
        public int Year;
        public string Term;
        public string SubjectArea;
        public int Semester;
        public DataScraperInput(int year, string subjectArea, int semester)
        {
            Year = year;
            Term = getTerm(semester);
            SubjectArea = subjectArea;
            Semester = semester;
        }

        string getTerm(int semester)
        {
            string termString;
            if (semester.Equals(1))
            {
                termString = "5" + Year.ToString().Last() + "40";
            }
            else
            {
                termString = "5" + Year.ToString().Last() + "80";
            }

            return termString;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker
{
    public class SubjectEnrollment
    {
        public String CourseCode { get; }
        public int Year { get; }
        public int Semester { get; }

        public SubjectEnrollment(string courseCode, int year, int semester)
        {
            this.CourseCode = courseCode;
            this.Year = year;
            this.Semester = semester;
        }
    }
}

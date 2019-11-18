using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace CoursePathwayMaker.NuStarDataScraperTool
{
    public class ClassRosterSearchResult
    {
        public string Term { get; private set; }
        public string SubjectArea { get; private set; }
        public string CourseCode { get; private set; }
        public List<Student> StudentsEnrolled { get; set; }
        public string CourseDescription { get; set; }
        public string Section { get; set; }
        public string Campus { get; set; }
        public string Year { get; set; }
        public string Semester { get; set; }

        public ClassRosterSearchResult(IWebElement searchResult, string year, string semester)
        {
            Year = year;
            Semester = semester;
            FillOutResultInfo(searchResult);
        }

        void FillOutResultInfo(IWebElement searchResult)
        {
            var info = searchResult.FindElements(By.TagName("a"));

            if (info.Any())
            {
                Term = info[1].Text;
                SubjectArea = info[2].Text;
                CourseCode = info[2].Text.Trim() + info[3].Text.Trim();
                CourseDescription = info[9].Text;
                Section = info[5].Text;
            }
        }

        void SetCampus()
        {
            // DONT FORGET THIS
        }
    }
}

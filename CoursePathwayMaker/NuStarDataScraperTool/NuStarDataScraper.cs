using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace CoursePathwayMaker.NuStarDataScraperTool
{
    public class NuStarDataScraper
    {
        public NuStarDataScraper()
        {
        }

        public void GetDataFromNuStarWebsite()
        {
            using (var driver = new ChromeDriver(@"C:\Users\kh462\Documents\DEV\CoursePathwayMaker\packages\Selenium.WebDriver.ChromeDriver.78.0.3904.7000\driver\win32\"))
            {
                Console.WriteLine("Scraper navigating to NuStar website...");

                Console.Write("Subject Area: ");
                var subjectArea = Console.ReadLine();

                Console.Write("Term: ");
                var term = Console.ReadLine();

                Console.WriteLine("Info For Search Result...");
                Console.Write("Year: ");
                var year = Console.ReadLine();

                Console.Write("Semester: ");
                var semester = Console.ReadLine();

                var navigator = new NuStarWebsiteNavigator(driver);
                navigator.NavigateToUonWebsite();
                navigator.FillSearchFilters(subjectArea, term);
                var results = GetSearchResults(driver, navigator, year, semester);
                SaveResultsInExcelWorksheet(results, year, semester, term, subjectArea);
            }
        }

        public List<ClassRosterSearchResult> GetSearchResults(ChromeDriver driver, NuStarWebsiteNavigator navigator, string year, string semester)
        {
            driver.FindElement(By.Id("CLASS_ROSTER")).Submit();
            var table = driver.FindElementById("PTSRCHRESULTS");
            var tableRows = table.FindElements(By.XPath(".//tbody/tr"));

            List<ClassRosterSearchResult> results = new List<ClassRosterSearchResult>();

            var rowCount = 0;
            while (rowCount < tableRows.Count)
            {
                if (rowCount > 0)
                {
                    driver.FindElement(By.Id("CLASS_ROSTER")).Submit();
                    var row = driver.FindElementById("PTSRCHRESULTS").FindElements(By.XPath(".//tbody/tr"))[rowCount];
                    Console.WriteLine(row.FindElements(By.TagName("td")).Last().Text);
                    var newResult = new ClassRosterSearchResult(row, year, semester);

                    navigator.GoToSearchResultEnrollmentPage(row);

                    newResult.StudentsEnrolled = GetEnrollments(driver);
                    newResult.Campus = GetCampus(driver);

                    results.Add(newResult);
                    
                    navigator.ReturnToClassRosterSearch();
                    
                }
                rowCount++;
            }

            return results;
        }

        List<Student> GetEnrollments(ChromeDriver driver)
        {
            driver.FindElementById("CLASS_ROSTER").Submit();
            var enrollmentTable = driver.FindElements(By.ClassName("PSLEVEL1GRID"));

            if (enrollmentTable.Count > 1)
            {
                var tableRows = enrollmentTable[1].FindElements(By.XPath(".//tbody//tr"));

                List<Student> enrollments = new List<Student>();

                var rowCount = 0;
                foreach (var row in tableRows)
                {
                    if (rowCount > 0)
                    {
                        var enrollmentInfo = row.FindElements(By.TagName("td"));
                        var student = new Student(Convert.ToInt32(enrollmentInfo[1].Text));
                        student.ProgramAndPlan = enrollmentInfo[6].Text;
                        student.Level = enrollmentInfo[7].Text;

                        enrollments.Add(student);
                    }
                    rowCount++;
                }
                return enrollments;
            }
            else
            {
                return new List<Student>();
            }
            
        }

        string GetCampus(ChromeDriver driver)
        {
            var foundCampus = false;
            string campus = "Unknown";
            try
            {
                var label = driver.FindElementById("DERIVED_SSR_FC_SSR_CLASSNAME_LONG").Text;
                // var campusOption1 = Regex.Match(label, @"(?:.*- )(.*)(?: \(.*)");

                var roomLabel = driver.FindElementById("win0divMTG_LOC$0").Text;
                //var campusOption2 = Regex.Match(roomLabel, "(?:.* - )(.*)");
                foundCampus = true;

                if (label.Contains("OUR") || roomLabel.Contains("Ourimbah"))
                {
                    campus = "Ourimbah";
                }
                else if (label.Contains("NCL") || roomLabel.Contains("Newcastle City Precinct"))
                {
                    campus = "City";
                }
                else if (label.Contains("CAL") || roomLabel.Contains("Callaghan"))
                {
                    campus = "Callaghan";
                }
            }
            catch { }

            return campus;
        }

        void SaveResultsInExcelWorksheet(List<ClassRosterSearchResult> results, string year, string semester, string term, string subjectArea)
        {
            Application App = new Application();
            try
            {
                Console.WriteLine("Populating Spreadsheet...");
                var name = year + "-" + semester + "-" + term + "-" + subjectArea;
                Workbook workbook = App.Workbooks.Add();
                Worksheet worksheet = workbook.Worksheets.Add() as Worksheet;
                worksheet.Name = name;

                worksheet.Cells[1, 1] = "Year";
                worksheet.Cells[1, 2] = "Semester";
                worksheet.Cells[1, 3] = "Course Code";
                worksheet.Cells[1, 4] = "ID";
                worksheet.Cells[1, 5] = "Campus";
                worksheet.Cells[1, 6] = "SubjectArea";
                worksheet.Cells[1, 7] = "Course Description";
                worksheet.Cells[1, 8] = "Section";
                worksheet.Cells[1, 9] = "Student Program and Plan";

                var rowNum = 2;
                foreach (var result in results)
                {
                    foreach (var enrollment in result.StudentsEnrolled)
                    {
                        worksheet.Cells[rowNum, 1] = result.Year;
                        worksheet.Cells[rowNum, 2] = Convert.ToInt32(result.Semester);
                        worksheet.Cells[rowNum, 3] = result.CourseCode;
                        worksheet.Cells[rowNum, 4] = enrollment.StudentID;
                        worksheet.Cells[rowNum, 5] = result.Campus;
                        worksheet.Cells[rowNum, 6] = result.SubjectArea;
                        worksheet.Cells[rowNum, 7] = result.CourseDescription;
                        worksheet.Cells[rowNum, 8] = result.Section;
                        worksheet.Cells[rowNum, 9] = enrollment.ProgramAndPlan;

                        rowNum++;
                    }
                }

                var saveFileName = "C:\\Users\\kh462\\Documents\\grabbing info test\\" + name + ".xlsx";
                Console.WriteLine("Saving as: '{0}'", saveFileName);
                workbook.SaveAs(saveFileName);
                App.Quit();
                Console.WriteLine("COMPLETE :)");
            }
            catch (Exception ex)
            {
                App.Quit();
                throw;
            }
        }
    }
}

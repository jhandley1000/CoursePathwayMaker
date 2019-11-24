using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CoursePathwayMaker.PathwayMaker;
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

        public void GetDataFromNuStarWebsite(IConsoleReader consoleReader, ISaveFilePathMaker saveFilePathMaker)
        {
            var userInputs = GetInputFromUser(consoleReader);
            var dbOrExcel = consoleReader.AddToDb();

            using (var driver = new ChromeDriver(@"E:\CoursePathwayMaker\CoursePathwayMaker\bin\Debug"))
            {
                Console.WriteLine("Scraper navigating to NuStar website...");


                //var subjectArea = consoleReader.GetSubjectArea();
                //var term = consoleReader.GetTerm();
                //var year = consoleReader.GetStartYear().ToString();
                //var semester = consoleReader.GetSemester();

                var navigator = new NuStarWebsiteNavigator(driver);
                navigator.NavigateToUonWebsite();
                navigator.LoginIfNecessary();
                navigator.NavigateToUonWebsiteAgain();
                List<ClassRosterSearchResult> results = new List<ClassRosterSearchResult>();
                foreach (var sem in userInputs)
                {
                    navigator.FillSearchFilters(sem.SubjectArea, sem.Term);
                    results.AddRange(GetSearchResults(driver, navigator, sem.Year.ToString(), sem.Semester.ToString()));
                    navigator.ClearSearchFields();
                }

                if (dbOrExcel == true)
                {
                    SaveResultsInDb(results);
                }
                else
                {
                    SaveResultsInExcelWorksheet(results, saveFilePathMaker);
                }
                
            }
        }

        List<DataScraperInput> GetInputFromUser(IConsoleReader consoleReader)
        {
            Console.WriteLine("Info For Search Result...");

            var subjectArea = consoleReader.GetSubjectArea();
            var startYear = consoleReader.GetStartYear();
            var endYear = consoleReader.GetEndYear();

            var dataScraperInput = new List<DataScraperInput>();

            while (startYear <= endYear)
            {
                dataScraperInput.Add(new DataScraperInput(startYear, subjectArea, 1));
                dataScraperInput.Add(new DataScraperInput(startYear, subjectArea, 2));
                startYear++;
            }

            return dataScraperInput;
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

        void SaveResultsInDb(List<ClassRosterSearchResult> results)
        {
            var dbHandler = new DbHandler();
            try
            {
                var queryStrings = new List<string>();
                dbHandler.OpenConnection();
                var queryNum = 0;
                var index = 1;
                foreach (var result in results)
                {
                    foreach (var enrollment in result.StudentsEnrolled)
                    {
                        if (queryNum < 1000)
                        {
                            dbHandler.AddNuStarDataRowToInsertQuery(Convert.ToInt32(result.Year), Convert.ToInt32(result.Semester), result.CourseCode, enrollment.StudentID, result.Campus, result.SubjectArea, result.CourseDescription, result.Section, enrollment.ProgramAndPlan);
                        } else
                        {
                            queryStrings.Add(dbHandler.sql);
                            dbHandler.ReturnAndResetSqlString();
                            dbHandler.AddNuStarDataRowToInsertQuery(Convert.ToInt32(result.Year), Convert.ToInt32(result.Year), result.CourseCode, enrollment.StudentID, result.Campus, result.SubjectArea, result.CourseDescription, result.Section, enrollment.ProgramAndPlan);
                            queryNum = 0;
                        }

                        queryNum += 1;
                        index += 1;
                    }
                }

                queryStrings.Add(dbHandler.sql);

                foreach (var query in queryStrings)
                {
                    dbHandler.AddNuStarDataToDb(query);
                }
                dbHandler.CloseConnection();

            }
            catch (DbException ex)
            {
               dbHandler.CloseConnection();
            }
        }

        void SaveResultsInExcelWorksheet(List<ClassRosterSearchResult> results, ISaveFilePathMaker saveFilePathMaker)
        {
            Application App = new Application();
            try
            {
                Console.WriteLine("Populating Spreadsheet...");
                var name = results[0].Year + "-" + results[0].Semester + "-" + results[0].Term + "-" + results[0].SubjectArea;
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

                var saveFileName =saveFilePathMaker.Make(name);
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

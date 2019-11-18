using System;
using UnitTestProject1.TestObjects;
using CoursePathwayMaker.PathwayMaker;
using Microsoft.CSharp.RuntimeBinder;
using CoursePathwayMaker;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.IO;
using NUnit.Framework;
using System.Linq;

namespace UnitTestProject1
{
    public class PathwayMakerTest
    {
        [Test]
        public void TestOurimbahMNGT30072018S1Pathways_OnlyOurimbahData()
        {
            GeneratePathwaysForListOfStudentsFromEnrollmentData("OurimbahOnlyMNGTData2017-2019.xlsx", "OurimbahMNGT30072018S1TESTFILE.xlsx", "Ourimbah", 2017, 2019);
        }

        [Test]
        public void TestOurimbahMNGT30072018S1Pathways_OurimbahAndNCLData()
        {
            GeneratePathwaysForListOfStudentsFromEnrollmentData("AllCampusACFIData2016-2019", "AllCampusACFI10012016S1TESTFILE", "AllCampus", 2016, 2019);
        }

        public void GeneratePathwaysForListOfStudentsFromEnrollmentData(string dataFile, string testFilePath, string worksheetName, int startYear, int endYear)
        {

            var testExelDataFilePath = string.Format("{0}..\\..\\TestDataFile\\{1}.xlsx", System.AppDomain.CurrentDomain.BaseDirectory, dataFile);
            var testOutputFilePath = string.Format("{0}..\\..\\TestExcelFilesOUTPUT\\{1}_OUTPUT.xlsx", AppDomain.CurrentDomain.BaseDirectory, testFilePath);

            var consoleReader = new ConsoleReaderForTests(testExelDataFilePath,
                                                            startYear,
                                                            endYear,
                                                            worksheetName,
                                                            testOutputFilePath);
            var pathwayMakerTool = new PathwayMakerTool();

            try
            {
                pathwayMakerTool.MakePathways(consoleReader);

            }
            catch
            {
                throw;
            }
           
            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            var app = new Application();

            var outputWorkbook = (app.Workbooks.Open(testOutputFilePath));
            var outputFile = outputWorkbook.Worksheets.get_Item(worksheetName) as Worksheet;
            outputFile.Activate();

            var testFile = app.Workbooks.Open(string.Format("{0}..\\..\\TestExcelFiles\\{1}.xlsx", AppDomain.CurrentDomain.BaseDirectory, testFilePath)).Worksheets.get_Item(worksheetName);
            (testFile as Worksheet).Activate();
            Assert.AreEqual((testFile as Worksheet).UsedRange.Rows.Count, (outputFile as Worksheet).UsedRange.Rows.Count);
            Assert.AreEqual((testFile as Worksheet).UsedRange.Columns.Count, (outputFile as Worksheet).UsedRange.Columns.Count);
                
            try
            {

                var rowCount = 2;
                while (rowCount <= (testFile as Worksheet).UsedRange.Rows.Count)
                {
                    var colCount = 2;
                    while (colCount <= (testFile as Worksheet).UsedRange.Columns.Count)
                    {
                        Assert.AreEqual((((testFile as Worksheet).Cells[rowCount, colCount] as Range).Value as string)?.Trim(), 
                            (((outputFile as Worksheet).Cells[rowCount, colCount] as Range).Value as string)?.Trim(), 
                            String.Format("Pathways didn't match. Student: {0} Semester {1}, {2}", 
                                    ((testFile as Worksheet).Cells[rowCount, 1] as Range).Value.ToString(), 
                                    ((testFile as Worksheet).Cells[2, colCount] as Range).Value.ToString(), 
                                    ((testFile as Worksheet).Cells[1, colCount] as Range).Value.ToString()));
                        colCount++;
                    }
                    rowCount++;
                }

                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outputFile);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outputWorkbook);

                System.IO.File.Delete(testOutputFilePath);

            } 
            catch (Exception ex)
            {
                outputWorkbook.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outputFile);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outputWorkbook);
                System.IO.File.Delete(testOutputFilePath);
                throw;
            }
        }
    }
}

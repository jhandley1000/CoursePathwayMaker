using System;
using TestPathwayMaker.TestObjects;
using CoursePathwayMaker.PathwayMaker;
using Microsoft.CSharp.RuntimeBinder;
using CoursePathwayMaker;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.IO;
using NUnit.Framework;
using System.Linq;

namespace TestPathwayMaker
{
    public class PathwayMakerTest
    {
        [Test]
        public void TestOurimbahMNGT30072018S1Pathways_OnlyOurimbahData()
        {
            var consoleReader = new ConsoleReaderForTests("OurimbahOnlyMNGTData2017-2019",
                                                            2017,
                                                            2019,
                                                            "Ourimbah",
                                                            "OurimbahMNGT30072018S1TESTFILE");

            GeneratePathwaysForListOfStudentsFromEnrollmentData(consoleReader);
        }

        [Test]
        public void TestOurimbahMNGT30072018S1Pathways_OurimbahAndNCLData()
        {
            var consoleReader = new ConsoleReaderForTests("AllCampusACFIData2016-2019",
                                                            2016,
                                                            2019,
                                                            "AllCampus",
                                                            "AllCampusACFI10012016S1TESTFILE");

            GeneratePathwaysForListOfStudentsFromEnrollmentData(consoleReader);
        }

        public void GeneratePathwaysForListOfStudentsFromEnrollmentData(ConsoleReaderForTests consoleReader)
        {   
            var pathwayMakerTool = new PathwayMakerTool();

            pathwayMakerTool.MakePathways(consoleReader);
           
             var tableComparer = new WorksheetTableComparerForTests(consoleReader.GetTestFilePath(), consoleReader.GetNewSaveFilePath(), consoleReader.GetWorksheetName());
             tableComparer.Compare();
                //var rowCount = 2;
                //while (rowCount <= (testFile as Worksheet).UsedRange.Rows.Count)
                //{
                //    var colCount = 2;
                //    while (colCount <= (testFile as Worksheet).UsedRange.Columns.Count)
                //    {
                //        Assert.AreEqual((((testFile as Worksheet).Cells[rowCount, colCount] as Range).Value as string)?.Trim(), 
                //            (((outputFile as Worksheet).Cells[rowCount, colCount] as Range).Value as string)?.Trim(), 
                //            String.Format("Pathways didn't match. Student: {0} Semester {1}, {2}", 
                //                    ((testFile as Worksheet).Cells[rowCount, 1] as Range).Value.ToString(), 
                //                    ((testFile as Worksheet).Cells[2, colCount] as Range).Value.ToString(), 
                //                    ((testFile as Worksheet).Cells[1, colCount] as Range).Value.ToString()));
                //        colCount++;
                //    }
                //    rowCount++;
                //}

              

            //} 
            //catch (Exception ex)
            //{
            //    outputWorkbook.Close();
            //    app.Quit();
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(outputFile);
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(outputWorkbook);
            //    System.IO.File.Delete(consoleReader.GetNewSaveFilePath());
            //    throw;
            //}
        }
    }
}

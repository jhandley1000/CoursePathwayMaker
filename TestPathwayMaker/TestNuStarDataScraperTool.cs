using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestPathwayMaker.TestObjects;
using NUnit;
using NUnit.Framework;
using CoursePathwayMaker.NuStarDataScraperTool;
using UnitTestProject1.TestObjects;

namespace TestNuStarDataScraper
{
    public class TestNuStarDataScraperTool
    {
        [Test]
        public void GetDataFromNuStarWebsiteTest_MNGTS12019()
        {

            var consoleReader = new ConsoleReaderForTests("MNGT", "1", 2019, "5940");

            var nuStarDataScraper = new NuStarDataScraper();
            nuStarDataScraper.GetDataFromNuStarWebsite(consoleReader, new SaveFilePathMakerForTests());

            var filePathConstructor = new FilePathConstructorForTests();

            var tableComparer = new WorksheetTableComparerForTests(
                    filePathConstructor.ConstructExcelFilePath("TestExcelFiles", "GetDataFromNuStarWebsiteTest_MNGTS12019"), 
                    filePathConstructor.ConstructExcelFilePath("OUTPUTFORTEST", "2019-1-5940-MNGT"), 
                    "2019-1-5940-MNGT");

            tableComparer.Compare();
        }
    }
}

using NUnit.Framework;
using TestPathwayMaker.TestObjects;
using CoursePathwayMaker.ForceDirectedTableMaker;
using Microsoft.Office.Interop.Excel;

namespace ForceDirectedTableMakerTests
{
    public class TestForceDirectedTableMaker
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void CreateForceDirectedTable_OurimbahMNGT20062018S2()
        {
            var consoleReader = new ConsoleReaderForTests("OurimbahMNGT20062018S2PATHWAY", "OurimbahMNGT20062018S2OUTPUT", "OurimbahMNGT20062018S1TESTFILE1");

            var forceDirectedTableMaker = new ForceDirectedTableMakerTool();
            forceDirectedTableMaker.CreateForceDirectedTable(consoleReader);

			var worksheetTableComparer = new WorksheetTableComparerForTests(consoleReader.GetTestFilePath(), consoleReader.GetNewSaveFilePath(), "ForceDirectedTable");
			worksheetTableComparer.Compare();
        }
    }
}
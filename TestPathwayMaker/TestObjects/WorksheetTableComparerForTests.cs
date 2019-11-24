using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;

namespace TestPathwayMaker.TestObjects
{
    public class WorksheetTableComparerForTests
    {
		Application app;
		Workbook outputWorkbook;
        Worksheet testWorksheet { get; }
        Worksheet generatedWorksheet { get; }
		string generatedWorksheetPath { get; }
        ConsoleReaderForTests consoleReader { get; }

        public WorksheetTableComparerForTests(string testWorksheet, string generatedWorksheet, string worksheetName)
        {
			this.app = new Application();

			System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

			this.generatedWorksheetPath = generatedWorksheet;
			this.outputWorkbook = app.Workbooks.Open(testWorksheet);
			this.testWorksheet = outputWorkbook.Worksheets.get_Item(worksheetName) as Worksheet;
			this.testWorksheet.Activate();
            this.generatedWorksheet = app.Workbooks.Open(generatedWorksheet).Worksheets[worksheetName] as Worksheet;
		}

        public void Compare()
        {
			//Assert.DoesNotThrow(() => (testWorksheet.Cells[1, 1])
			try
			{


				Assert.AreEqual((testWorksheet as Worksheet).UsedRange.Rows.Count, (generatedWorksheet as Worksheet).UsedRange.Rows.Count);
				Assert.AreEqual((testWorksheet as Worksheet).UsedRange.Columns.Count, (generatedWorksheet as Worksheet).UsedRange.Columns.Count);

				var rowCount = 1;
				while (rowCount <= (testWorksheet as Worksheet).UsedRange.Rows.Count)
				{
					var colCount = 1;
					while (colCount <= (testWorksheet as Worksheet).UsedRange.Columns.Count)
					{
						var test = (testWorksheet.Cells[rowCount, colCount] as Range).Value;
						Assert.That(() => (testWorksheet.Cells[rowCount, colCount] as Range).Value, Throws.Nothing);
						Assert.AreEqual((((testWorksheet as Worksheet).Cells[rowCount, colCount] as Range)?.Value as string)?.Trim(),
							(((generatedWorksheet as Worksheet).Cells[rowCount, colCount] as Range)?.Value as string)?.Trim(),
							String.Format("Pathways didn't match. Student: {0} Semester {1}, {2}",
									((testWorksheet as Worksheet).Cells[rowCount, 1] as Range)?.Value?.ToString(),
									((testWorksheet as Worksheet).Cells[2, colCount] as Range)?.Value?.ToString(),
									((testWorksheet as Worksheet).Cells[1, colCount] as Range)?.Value?.ToString()));
						colCount++;
					}
					rowCount++;
				}

				QuitAndDeleteOutputFile();
			}
			catch
			{

				QuitAndDeleteOutputFile();
				throw;
			}
        }

		void QuitAndDeleteOutputFile()
		{
			app.Quit();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(generatedWorksheet);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(outputWorkbook);

			System.IO.File.Delete(generatedWorksheetPath);
		}
    }
}

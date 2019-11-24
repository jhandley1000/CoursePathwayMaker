using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoursePathwayMaker.PathwayMaker;

namespace CoursePathwayMaker.ForceDirectedTableMaker
{
    public class ForceDirectedTableMakerTool
    {
        public ForceDirectedTableMakerTool()
        {
        }

        public void CreateForceDirectedTable(IConsoleReader consoleReader)
        {
            var pathwayTableFilePath = consoleReader.GetInputFilePath();
            var forceDirectedTableFilePath = consoleReader.GetNewSaveFilePath();

            var handler = new ForceDirectedTableExcelHandler(pathwayTableFilePath, forceDirectedTableFilePath);
            handler.GetConnectionsFromPathwayFile();
			handler.PutAllConnectionsInOutputFile();
			handler.SaveForceDirectedWorkbook();

            handler.QuitApp();
        }
    }
}

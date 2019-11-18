using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.ForceDirectedTableMaker
{
    public class FDTableMaker
    {
        public FDTableMaker()
        {
        }

        public void CreateForceDirectedTable()
        {
            Console.Write("File Path for Pathway Table: ");
            var pathwayTableFilePath = Console.ReadLine();

            Console.Write("Save As: ");
            var forceDirectedTableFilePath = Console.ReadLine();

            var handler = new ForceDirectedTableExcelHandler(pathwayTableFilePath, forceDirectedTableFilePath);
            handler.GetConnectionsFromPathwayFile();
			handler.PutAllConnectionsInOutputFile();
			handler.SaveForceDirectedWorkbook();

            handler.QuitApp();
        }
    }
}

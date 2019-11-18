using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoursePathwayMaker.ForceDirectedTableMaker;
using CoursePathwayMaker.NuStarDataScraperTool;
using CoursePathwayMaker.PathwayMaker;

namespace CoursePathwayMaker
{
	class Program
	{
		static void Main(string[] args)
		{
            Console.WriteLine(@"Welcome to Jess' pathway generator. This is still in dev, so if it breaks let me know.
Commands:
mp      ------ make pathway table from enrollmentdata file
prepfd  ------ prepare data for force directed diagram from pathway table
getdata ------ get data from NuStar");

            string command;
            while ((command = Console.ReadLine()) != "q")
            {
                if (command.Equals("mp"))
                {
                    var pathwayMaker = new PathwayMakerTool();
                    pathwayMaker.MakePathways(new ConsoleReader());
                }
                
                if (command.Equals("prepfd"))
                {
                    var forceDirectedTableMaker = new FDTableMaker();
                    forceDirectedTableMaker.CreateForceDirectedTable();
                }

                if (command.Equals("getdata"))
                {
                    var nuStarDataScraper = new NuStarDataScraper();
                    nuStarDataScraper.GetDataFromNuStarWebsite();
                }
            }
		}
	}
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoursePathwayMaker.ForceDirectedTableMaker;
using CoursePathwayMaker.NuStarDataScraperTool;
using CoursePathwayMaker.PathwayMaker;
using System.Data.Common;
using System.Configuration;

namespace CoursePathwayMaker
{
	class Program
	{
		static void Main(string[] args)
		{
            //string provider = ConfigurationManager.AppSettings["provider"];
            //string connectionString = ConfigurationManager.AppSettings["connectionString"];

            //DbProviderFactory factory = DbProviderFactories.GetFactory(provider);

            //using (DbConnection connection = factory.CreateConnection())
            //{
            //    if (connection == null)
            //    {
            //        Console.WriteLine("Connection Error");
            //        Console.ReadLine();
            //        return;
            //    }

            //    connection.ConnectionString = connectionString;

            //    connection.Open();

            //    DbCommand command1 = factory.CreateCommand();

            //    if (command1 == null)
            //    {
            //        Console.WriteLine("Command Error");
            //        Console.ReadLine();
            //        return;
            //    }

            //    command1.Connection = connection;

            //    command1.CommandText = "Select * From SubjectEnrollments";

            //    using (DbDataReader dataReader = command1.ExecuteReader())
            //    {
            //        while (dataReader.Read())
            //        {
            //            Console.WriteLine($"{dataReader["CourseCode"]} " + $"{dataReader["StudentID"]}");
            //        }
            //    }
            //}

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
                    var forceDirectedTableMaker = new ForceDirectedTableMakerTool();
                    forceDirectedTableMaker.CreateForceDirectedTable(new ConsoleReader());
                }

                if (command.Equals("getdata"))
                {
                    var nuStarDataScraper = new NuStarDataScraper();
                    nuStarDataScraper.GetDataFromNuStarWebsite(new ConsoleReader(), new SaveFilePathMaker());
                }
            }
		}
	}
}

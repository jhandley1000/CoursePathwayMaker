using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.NuStarDataScraperTool
{
    public class SaveFilePathMaker : ISaveFilePathMaker
    {
        public string Make(string name)
        {
            return "C:\\Users\\kh462\\Documents\\grabbing info test\\" + name + ".xlsx";
        }
    }
}

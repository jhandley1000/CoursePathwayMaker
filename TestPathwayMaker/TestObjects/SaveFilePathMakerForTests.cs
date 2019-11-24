using CoursePathwayMaker.NuStarDataScraperTool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestPathwayMaker.TestObjects;

namespace UnitTestProject1.TestObjects
{
    public class SaveFilePathMakerForTests : ISaveFilePathMaker
    {
        public string Make(string name)
        {
            return new FilePathConstructorForTests().ConstructExcelFilePath("OUTPUTFORTEST", name);
        }
    }
}

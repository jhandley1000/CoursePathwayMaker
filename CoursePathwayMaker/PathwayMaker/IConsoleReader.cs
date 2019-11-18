using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.PathwayMaker
{
    public interface IConsoleReader
    {
        string GetDataFilePath();
        int GetStartYear();
        int GetEndYear();
        string GetCampus();
        string GetFileSavePath(); 
    }
}

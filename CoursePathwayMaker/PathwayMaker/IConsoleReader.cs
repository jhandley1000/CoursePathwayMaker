using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.PathwayMaker
{
    public interface IConsoleReader
    {
        string GetInputFilePath();
        int GetStartYear();
        int GetEndYear();
        string GetWorksheetName();
        string GetNewSaveFilePath();
        string GetSubjectArea();
        string GetTerm();
        string GetSemester();
        bool AddToDb();
    }
}

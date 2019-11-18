using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker.ForceDirectedTableMaker
{
    public class Connection
    {
        public string FromCourseCode { get; }
        public string ToCourseCode { get; }
        public int PathwayFrequency { get; private set; }

        public Connection(string fromCourseCode, string toCourseCode)
        {
            FromCourseCode = fromCourseCode;
            ToCourseCode = toCourseCode;
            PathwayFrequency = 1;
        }

        public void AddToPathwayFrequency()
        {
            PathwayFrequency++;
        }
    }
}

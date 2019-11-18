using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoursePathwayMaker
{
    public class Student
    {
        public int StudentID { get; }
        public List<SubjectEnrollment>  SubjectEnrollments { get; }
        public string ProgramAndPlan { get; set; }
        public string Level { get; set; }

        public Student(int studentID)
        {
            this.SubjectEnrollments = new List<SubjectEnrollment>();
            this.StudentID = studentID;
        }

        public void AddSubjectEnrollment(SubjectEnrollment subjectEnrollment)
        {
            SubjectEnrollments.Add(subjectEnrollment);
        }
    }
}

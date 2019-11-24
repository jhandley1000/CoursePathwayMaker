using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace CoursePathwayMaker
{
	public class DbHandler
	{
		string provider = ConfigurationManager.AppSettings["provider"];
		string connectionString = ConfigurationManager.AppSettings["homeConnectionString"];
		DbProviderFactory factory;
        SqlDataAdapter dataAdapter;
        public string sql = "INSERT INTO SubjectEnrollments (Id, Year, Semester, CourseCode, StudentID, Campus, SubjectArea, CourseDescription, Section, ProgramAndPlan) VALUES ";
        DbConnection connection;
        int rowsAdded;
		public DbHandler()
		{
			factory = DbProviderFactories.GetFactory(provider);
            dataAdapter = new SqlDataAdapter();
            rowsAdded = 0;
		}

		public DbConnection OpenConnection()
		{
            connection = factory.CreateConnection();
			{
				if (connection == null)
				{
					Console.WriteLine("Connection Error");
					Console.ReadLine();
					throw new Exception("Connection Error");
				}

				connection.ConnectionString = connectionString;

				connection.Open();

				return connection;
			}
		}
		
		public List<Student> GetStudentEnrollments(List<int> studentIDs)
		{
			var studentEnrollments = new List<Student>();
			using (var connection = OpenConnection())
			{
				foreach (var student in studentIDs)
				{
					studentEnrollments.Add(PopulateStudentWithEnrollment(student, connection));
				}
			}

			return studentEnrollments;
		}

		Student PopulateStudentWithEnrollment(int student, DbConnection connection)
		{
			Student newStudent = new Student(student);

			var command = factory.CreateCommand();
			command.Connection = connection;
			command.CommandText = string.Format("SELECT * FROM ENROLLMENT WHERE STUDENTID = {0}", student);

			using (DbDataReader dataReader = command.ExecuteReader())
			{
				while (dataReader.Read())
				{
					var enrollment = new SubjectEnrollment(
						dataReader["CourseCode"].ToString(),
						Convert.ToInt32(dataReader["Year"]),
						Convert.ToInt32(dataReader[""])
						);

					newStudent.AddSubjectEnrollment(enrollment);
				}
			}

			return newStudent;

		}

        public void AddNuStarDataRowToInsertQuery(int year, int semester, string courseCode, int studentID, string campus, string subjectArea, string courseDescription, string section, string programAndPlan)
        {
            sql += string.Format("(NEWID(), {0}, {1}, '{2}', {3}, '{4}', '{5}', '{6}', '{7}', '{8}'),", year, semester, courseCode, studentID, campus, subjectArea, courseDescription, section, programAndPlan);
            rowsAdded += 1;
        } 

        public string ReturnAndResetSqlString()
        {
            var returnSql = sql;
            sql = "INSERT INTO SubjectEnrollments (Id, Year, Semester, CourseCode, StudentID, Campus, SubjectArea, CourseDescription, Section, ProgramAndPlan) VALUES ";
            return returnSql;
        }

        public void AddNuStarDataToDb(string sqlString)
        {
            //var command = new SqlCommand(sql, connection);

            //dataAdapter.InsertCommand = new SqlCommand(sql, connection);
            //dataAdapter.InsertCommand.ExecuteNonQuery();

            //Console.WriteLine("{0} rows added successfully.", rowsAdded);
            //rowsAdded = 0;
            //command.Dispose();

            //DbProviderFactory factory = DbProviderFactories.GetFactory(provider);
            sqlString = sqlString.Remove(sqlString.Length - 1);
            using (DbConnection connection = factory.CreateConnection())
            {
                if (connection == null)
                {
                    Console.WriteLine("Connection Error");
                    Console.ReadLine();
                    return;
                }

                connection.ConnectionString = connectionString;

                connection.Open();

                using (DbCommand command1 = factory.CreateCommand())
                {
                    if (command1 == null)
                    {
                        Console.WriteLine("Command Error");
                        Console.ReadLine();
                        return;
                    }

                    command1.Connection = connection;

                    command1.CommandText = sqlString;

                    command1.ExecuteNonQuery();

                    Console.WriteLine("{0} rows added.", rowsAdded);
                    rowsAdded = 0;
                }
            }
        }

        public void CloseConnection()
        {
            connection.Close();
        }
        //DbProviderFactory factory = DbProviderFactories.GetFactory(provider);

        //			using (DbConnection connection = factory.CreateConnection())
        //			{
        //				if (connection == null)
        //				{
        //					Console.WriteLine("Connection Error");
        //					Console.ReadLine();
        //					return;
        //				}

        //connection.ConnectionString = connectionString;

        //				connection.Open();

        //	DbCommand command1 = factory.CreateCommand();

        //	if (command1 == null)
        //	{
        //		Console.WriteLine("Command Error");
        //		Console.ReadLine();
        //		return;
        //	}

        //	command1.Connection = connection;

        //	command1.CommandText = "Select * From Enrollment";

        //	using (DbDataReader dataReader = command1.ExecuteReader())
        //	{
        //		while (dataReader.Read())
        //		{
        //			Console.WriteLine($"{dataReader["CourseCode"]} " + $"{dataReader["StudentID"]}");
        //		}
        //	}
        //}

    }
}

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EasyWordDoc_DBResoration
{
    public class WordProcessRepo
    {
        static string ConnectionString = "Data Source=.;Initial Catalog=WordProcessNew;Integrated Security=True";

        public static List<int> GetAllQuestionId(string testId)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select Qid, Question, Hasimage, Subject, SClass, QType, Marks, QDesc, Topic, SheetType from Questions where Qid in (select QID from TestInfo where TestId = '" + testId + "')";
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    QuestionModel
                }
            }
            return new List<int>();
        }
        public static List<string> GetAllTestIds()
        {
            List<string> toReturn = new List<string>();
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select Distinct TestId from TestInfo";
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    toReturn.Add(reader.GetString(0));
                }
            }
            return toReturn;
        }

    }
}

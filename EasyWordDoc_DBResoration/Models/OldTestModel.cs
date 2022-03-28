using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Models
{
    public class OldTestModel
    {
        public string TestID { get; set; }
        public int Qid { get; set; }

        public static List<string> GetAllTestIds()
        {
            List<string> toReturn = new List<string>();
            using (SqlConnection con = new SqlConnection("Data Source=.;Initial Catalog=WordProcess;Integrated Security=True"))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select Distinct TestId from TestInfo";
                SqlDataReader reader = cmd.ExecuteReader();
                while(reader.Read())
                {
                    toReturn.Add(reader.GetString(0));
                }
            }
            return toReturn;
        }
    }
}

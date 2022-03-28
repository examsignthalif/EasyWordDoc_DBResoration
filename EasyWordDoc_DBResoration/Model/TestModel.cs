using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Model
{
    public class TestModel
    {
        public string TestID { get; set; }
        public List<QuestionModel> QuestionList { get; set; }

        public TestModel()
        {
            QuestionList = new List<QuestionModel>();
        }
    }
}

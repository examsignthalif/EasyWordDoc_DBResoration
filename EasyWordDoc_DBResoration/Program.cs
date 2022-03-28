using EasyWordDoc_DBResoration.Model;
using EasyWordDoc_DBResoration.Repo;
using EasyWordDoc_DBResoration.Tools;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration
{
    class Program
    {
        static string ConnectionString = "Data Source=.;Initial Catalog=WordProcess;Integrated Security=True";
        static string ConnectionString2 = "Data Source=.;Initial Catalog=WordProcessNew;Integrated Security=True";
        static SqlConnection con1 = new SqlConnection(ConnectionString);
        static SqlConnection con2 = new SqlConnection(ConnectionString2);

        static List<TestModel> TestList = new List<TestModel>();
        static void Main(string[] args)
        {
            con1.Open();
            con2.Open();

            Process();
            Console.ReadKey();
        }

        static void Process()
        {
            // 1. Get all TestId for a class : 7
            // 2. Now get all Question for TestId list.

            // 3. Each TestId == FileId

            // 4. Insert Question
            // 5. Insert ImageTable
            // 6. Insert XpsTable
            // 7. Insert QuestionOrigin
            // 8. Modify Insert TestInfo
            // 9. Insert IsHeadingUpdate
            CollectQuestion(5);
            ReStoreProcess();
        }
        static void CollectQuestion(int givenGrade)
        {
            List<string> TestIdList = WordProcessRepo.GetAllTestIdForGrade(givenGrade);
            foreach (string testId in TestIdList)
            {
                TestModel obj = new TestModel();
                obj.TestID = testId;
                obj.QuestionList = WordProcessRepo.GetAllQuestionForTestId(testId);

                TestList.Add(obj);
            }
        }

        static void ReStoreProcess()
        {
            List<string> failedTestId = new List<string>();
            foreach(TestModel T_Item in TestList)
            {
                try
                {
                    foreach (QuestionModel Q_Item in T_Item.QuestionList)
                    {
                        WordProcessRepo.InsertQuestionItem(con2, Q_Item);
                        WordProcessRepo.InsertImagesItem(con2, Q_Item);
                        if (Q_Item.XpsByteData != null)
                            WordProcessRepo.InsertXpsItem(con2, Q_Item);
                        WordProcessRepo.InsertQuestionOrigin(con2, T_Item.TestID, Q_Item.Qid);
                        WordProcessRepo.InsertTestInfo(con2, T_Item.TestID, Q_Item.Grade, Q_Item.Subject, Q_Item.Qid);
                    }
                    WordProcessRepo.InsertIsHeadingUpdate(con2, T_Item.TestID, false);
                    Console.WriteLine("OK");
                }
                catch(Exception ex)
                {
                    failedTestId.Add(T_Item.TestID);
                    Console.WriteLine("FAILED:\t"+ ex.Message);
                }
            }

            Console.WriteLine("\n");
            foreach(string failedTest in failedTestId)
            {
                Console.WriteLine(failedTest);
            }
        }
    }
}

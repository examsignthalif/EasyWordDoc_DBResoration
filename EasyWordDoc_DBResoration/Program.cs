using EasyWordDoc_DBResoration.Model;
using EasyWordDoc_DBResoration.Repo;
using EasyWordDoc_DBResoration.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration
{
    class Program
    {
        static List<TestModel> TestList = new List<TestModel>();
        static void Main(string[] args)
        {
            
        }

        void Process()
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
        void CollectQuestion(int givenGrade)
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

        void ReStoreProcess()
        {
            List<string> failedTestId = new List<string>();
            foreach(TestModel T_Item in TestList)
            {
                try
                {
                    foreach (QuestionModel Q_Item in T_Item.QuestionList)
                    {
                        WordProcessRepo.InsertQuestionItem(Q_Item);
                        WordProcessRepo.InsertImagesItem(Q_Item);
                        if (Q_Item.XpsByteData.Length > 0)
                            WordProcessRepo.InsertXpsItem(Q_Item);
                        WordProcessRepo.InsertQuestionOrigin(T_Item.TestID, Q_Item.Qid);
                        WordProcessRepo.InsertTestInfo(T_Item.TestID, Q_Item.Grade, Q_Item.Subject, Q_Item.Qid);
                    }
                    WordProcessRepo.InsertIsHeadingUpdate(T_Item.TestID, false);
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

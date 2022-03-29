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
        static int Grade = 6;
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
            CollectQuestion(Grade);
            ReStoreProcess();
            Console.WriteLine("Completed..!");
        }
        static void CollectQuestion(int givenGrade)
        {
            List<string> TestIdList = WordProcessRepo.GetAllTestIdForGrade(givenGrade);
            foreach (string testId in TestIdList)
            {
                TestModel obj = new TestModel();
                obj.TestID = testId;
                obj.QuestionList = WordProcessRepo.GetAllQuestionForTestId(WordProcessRepo.GetQuestionIdList(testId));

                TestList.Add(obj);
            }
        }

        static void ReStoreProcess()
        {
            int NewTestId = WordProcessRepo.GetAllTestCount(con2);
            int NewQid = WordProcessRepo.GetAllQuestionCount(con2);
            foreach(TestModel T_Item in TestList)
            {
                try
                {
                    NewTestId++;
                    foreach (QuestionModel Q_Item in T_Item.QuestionList)
                    {
                        NewQid++;
                        try
                        { WordProcessRepo.InsertQuestionItem(con2, NewQid, Q_Item); }
                        catch (Exception ex) 
                        {
                        }
                        try
                        { WordProcessRepo.InsertImagesItem(con2, NewQid, Q_Item); }
                        catch (Exception ex) 
                        { 
                        }
                        if (Q_Item.XpsByteData != null)
                        {
                            try
                            { WordProcessRepo.InsertXpsItem(con2, NewQid, Q_Item); }
                            catch (Exception ex) 
                            {
                            }
                        }
                            
                        try
                        { WordProcessRepo.InsertQuestionOrigin(con2, NewTestId.ToString(), NewQid); }
                        catch (Exception ex) 
                        { 
                        }
                        try
                        { WordProcessRepo.InsertTestInfo(con2, NewTestId.ToString(), Q_Item.Grade, Q_Item.Subject, NewQid); }
                        catch (Exception ex) 
                        {
                        }
                    }
                    try
                    {
                        WordProcessRepo.InsertIsHeadingUpdate(con2, NewTestId.ToString(), false);
                    }
                    catch
                    {

                    }
                    Console.WriteLine("OK\t- "+ NewTestId.ToString());
                }
                catch(Exception ex)
                {
                    Console.WriteLine("FAILED:\t"+ ex.Message);
                }
            }
        }
    }
}

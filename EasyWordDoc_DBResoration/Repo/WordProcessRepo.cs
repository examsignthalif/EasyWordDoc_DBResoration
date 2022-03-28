using EasyWordDoc_DBResoration.Model;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Repo
{
    public class WordProcessRepo
    {
        static string ConnectionString = "Data Source=.;Initial Catalog=WordProcess;Integrated Security=True";

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
        public static int GetNewQuestionId()
        {
            int NewQuestionId = 0;
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "select count(*) from Questions";
                NewQuestionId = Convert.ToInt32(cmd.ExecuteScalar()) + 1;
            }
            return NewQuestionId;
        }
        public static int GetNewQuestionFileId()
        {
            int toReturn = 0;
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "select FileId from UploadedQuestionFile order by FileId desc";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    toReturn = dr.GetInt32(0);
                    break;
                }
            }
            return (toReturn + 1);
        }


        public static List<QuestionModel> GetAllQuestion(string testId)
        {
            List<QuestionModel> toReturn = new List<QuestionModel>();
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "select Qid, Question, Hasimage, Subject, SClass, QType, Marks, QDesc, Topic, SheetType from Questions where Qid in (select QID from TestInfo where TestId = '" + testId + "')";
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    QuestionModel obj = new QuestionModel();
                    obj.Qid = reader.GetInt32(0);
                    obj.Question = reader.GetString(1);
                    obj.HasImage = reader.GetBoolean(2);
                    obj.Subject= reader.GetString(3);
                    obj.Grade = reader.GetString(4);
                    obj.QType = reader.GetString(5);
                    obj.Mark = reader.GetInt32(6);
                    obj.Heading = reader.GetString(7);
                    obj.Topic = reader.GetString(8);
                    obj.SheetType = reader.GetString(9);
                    obj.ImageList = GetImages(obj.Qid);

                    toReturn.Add(obj);
                }
            }
            return toReturn;
        }
        private static List<ImageModel> GetImages(int qid)
        {
            List<ImageModel> toReturn = new List<ImageModel>();
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "select Qid, imagenumber, ImageByte from imagetable where Qid = " + qid + "";
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    ImageModel obj = new ImageModel();
                    obj.Qid = reader.GetInt32(0);
                    obj.ImageNumber = reader.GetString(1);
                    obj.ImageByteData = (byte[])reader.GetValue(2);
                    toReturn.Add(obj);
                }
            }
            return toReturn;
        }


        //New Database Insert
        public static void InsertQuestionItem(int newQuestionId, QuestionModel question)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "insert into Questions (Qid, Question, Hasimage, Subject, SClass, QType, Marks,QDesc, Topic, SubTopic, SheetType, QName) values(@Qid, @Question, @Hasimage, @Subject, @SClass, @QType, @Marks,@QDesc, @Topic, @SubTopic, @SheetType, @QName)";
                cmd.Parameters.AddWithValue("@Qid", newQuestionId);
                cmd.Parameters.AddWithValue("@Question", question.Question);
                cmd.Parameters.AddWithValue("@Hasimage", question.HasImage);
                cmd.Parameters.AddWithValue("@Subject", question.Subject);
                cmd.Parameters.AddWithValue("@SClass", question.Grade);
                cmd.Parameters.AddWithValue("@QType", question.QType);
                cmd.Parameters.AddWithValue("@QDesc", question.Heading);
                cmd.Parameters.AddWithValue("@Marks", question.Mark);
                cmd.Parameters.AddWithValue("@Topic", RemoveEndDot(question.Topic));
                cmd.Parameters.AddWithValue("@SheetType", question.SheetType);
                cmd.ExecuteNonQuery();
            }
        }
        public static void InsertImagesItem(int newQuestionId, QuestionModel question)
        {
            List<ImageModel> questionimage = question.ImageList.Where(temp => temp.Qid == question.Qid).ToList();

            foreach (ImageModel image in questionimage)
            {
                using (SqlConnection con = new SqlConnection(ConnectionString))
                {
                    SqlCommand cmd = new SqlCommand();
                    con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = "insert into imagetable values(@Qid,@imagenumber,@ImageByte)";
                    cmd.Parameters.AddWithValue("@Qid", newQuestionId);
                    cmd.Parameters.AddWithValue("@imagenumber", image.ImageNumber);
                    cmd.Parameters.AddWithValue("@ImageByte", image.ImageByteData);
                    cmd.ExecuteNonQuery();
                }
            }
        }
        public static void InsertXpsItem(int newQuestionId, byte[] xpsByteData)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "insert into Xpstable (Qid, XpsFile) values(@Qid, @XpsFile)";
                cmd.Parameters.AddWithValue("@Qid", newQuestionId);
                cmd.Parameters.AddWithValue("@XpsFile", xpsByteData);
                cmd.ExecuteNonQuery();
            }
        }
        public static bool InsertQuestionOrigin(int fileId, int qid)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "insert into QuestionOrigin(FileId, Qid) values (@FileId, @Qid)";
                cmd.Parameters.AddWithValue("@FileId", fileId);
                cmd.Parameters.AddWithValue("@Qid", qid);
                try
                {
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch (Exception ex) { return false; }
            }
        }
        public static void BackUpQuestionFile(int grade, string subject, string fileName, byte[] fileBytes)
        {
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "insert into UploadedQuestionFile (FileId, Grade, Subject, FileTittle, WordBytes) values (@FileId, @Grade, @Subject, @FileTittle, @WordBytes)";
                cmd.Parameters.AddWithValue("@FileId", GetNewQuestionFileId());
                cmd.Parameters.AddWithValue("@Grade", grade);
                cmd.Parameters.AddWithValue("@Subject", subject);
                cmd.Parameters.AddWithValue("@FileTittle", fileName);
                cmd.Parameters.AddWithValue("@WordBytes", fileBytes);
                cmd.ExecuteNonQuery();
            }
        }


        static string RemoveEndDot(string givenString)
        {
            if (givenString[givenString.Length - 1] == '.')
                return givenString.Substring(0, givenString.Length - 2).Trim();
            else
                return givenString.Trim();
        }
    }
}
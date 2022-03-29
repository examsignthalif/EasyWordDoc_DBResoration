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
        static string ConnectionString2 = "Data Source=.;Initial Catalog=WordProcessNew;Integrated Security=True";

        public static List<string> GetAllTestIdForGrade(int grade)
        {
            List<string> toReturn = new List<string>();
            using(SqlConnection con = new SqlConnection(ConnectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select Distinct TestId from TestInfo where QID in (select Qid from Questions where SClass = " + grade + ")";
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    toReturn.Add(reader.GetString(0));
                }
            }
            return toReturn;
        }
        public static int GetNewQuestionFileId()
        {
            int toReturn = 0;
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                SqlCommand cmd = new SqlCommand();

                cmd.CommandText = "select FileId from UploadedQuestionFile order by FileId desc";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    toReturn = dr.GetInt32(0);
                    break;
                }
                return (toReturn + 1);
            }
        }

        public static List<string> GetQuestionIdList(string testId)
        {
            List<string> toReturn = new List<string>();
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select distinct QID from TestInfo where TestId = '" + testId + "'";
                SqlDataReader dr = cmd.ExecuteReader();
                while(dr.Read())
                {
                    toReturn.Add(dr.GetString(0));
                }
            }
            return toReturn;
        }

        public static List<QuestionModel> GetAllQuestionForTestId(List<string> questionIdList)
        {
            List<QuestionModel> toReturn = new List<QuestionModel>();

            foreach(string qid in questionIdList)
            {
                using (SqlConnection con = new SqlConnection(ConnectionString))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "select Qid, Question, Hasimage, Subject, SClass, QType, Marks, QDesc, Topic, SheetType from Questions where Qid = " + qid + "";
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        QuestionModel obj = new QuestionModel();
                        obj.Qid = reader.GetInt32(0);
                        obj.Question = reader.GetString(1);
                        obj.HasImage = reader.GetBoolean(2);
                        obj.Subject = reader.GetString(3);
                        obj.Grade = reader.GetString(4);
                        obj.QType = reader.GetString(5);
                        obj.Mark = reader.GetInt32(6);
                        obj.Heading = reader.GetString(7);
                        obj.Topic = reader.GetString(8);
                        obj.SheetType = reader.GetString(9);
                        obj.ImageList = GetImages(obj.Qid);
                        obj.XpsByteData = GetXpsByteData(obj.Qid);
                        toReturn.Add(obj);
                    }
                }
            }
            return toReturn;
        }
        private static List<ImageModel> GetImages(int qid)
        {
            List<ImageModel> toReturn = new List<ImageModel>();
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
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
        static byte[] GetXpsByteData(int qid)
        {
            byte[] toReturn = new byte[0];
            using (SqlConnection con = new SqlConnection(ConnectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select XpsFile from Xpstable where Qid = " + qid + "";
                try
                {
                    toReturn = (byte[])cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {

                }
            }
            return toReturn;
        }



        public static int GetAllTestCount(SqlConnection gc)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "select COUNT(distinct TestId) from TestInfo";
            int count =(int)cmd.ExecuteScalar();
            return count;
        }
        public static int GetAllQuestionCount(SqlConnection gc)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "select count(*) from Questions";
            int count = (int)cmd.ExecuteScalar();
            return count;
        }

        //New Database Insert
        public static void InsertQuestionItem(SqlConnection gc, int newQid, QuestionModel question)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into Questions (Qid, Question, Hasimage, Subject, SClass, QType, Marks,QDesc, Topic, SubTopic, SheetType, QName, Qheading) values(@Qid, @Question, @Hasimage, @Subject, @SClass, @QType, @Marks,@QDesc, @Topic, @SubTopic, @SheetType, @QName, @Qheading)";
            cmd.Parameters.AddWithValue("@Qid", newQid);
            cmd.Parameters.AddWithValue("@Question", question.Question);
            cmd.Parameters.AddWithValue("@Hasimage", question.HasImage);
            cmd.Parameters.AddWithValue("@Subject", question.Subject);
            cmd.Parameters.AddWithValue("@SClass", question.Grade);
            cmd.Parameters.AddWithValue("@QType", question.QType);
            cmd.Parameters.AddWithValue("@QDesc", question.Heading);
            cmd.Parameters.AddWithValue("@Marks", question.Mark);
            cmd.Parameters.AddWithValue("@Topic", RemoveEndDot(question.Topic));
            cmd.Parameters.AddWithValue("@SheetType", question.SheetType);

            cmd.Parameters.AddWithValue("@SubTopic", string.Empty);
            cmd.Parameters.AddWithValue("@QName", string.Empty);
            cmd.Parameters.AddWithValue("@Qheading", string.Empty);
            cmd.ExecuteNonQuery();
        }
        public static void InsertImagesItem(SqlConnection gc, int newQid, QuestionModel question)
        {
            foreach (ImageModel image in question.ImageList)
            {
                SqlConnection con = gc;
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "insert into imagetable values(@Qid,@imagenumber,@ImageByte)";
                cmd.Parameters.AddWithValue("@Qid", newQid);
                cmd.Parameters.AddWithValue("@imagenumber", image.ImageNumber);
                cmd.Parameters.AddWithValue("@ImageByte", image.ImageByteData);
                cmd.ExecuteNonQuery();
            }
        }
        public static void InsertXpsItem(SqlConnection gc, int newQid, QuestionModel question)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into Xpstable (Qid, XpsFile) values(@Qid, @XpsFile)";
            cmd.Parameters.AddWithValue("@Qid", newQid);
            cmd.Parameters.AddWithValue("@XpsFile", question.XpsByteData);
            cmd.ExecuteNonQuery();
        }
        public static bool InsertQuestionOrigin(SqlConnection gc, string fileId, int newQid)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into QuestionOrigin(FileId, Qid) values (@FileId, @Qid)";
            cmd.Parameters.AddWithValue("@FileId", fileId);
            cmd.Parameters.AddWithValue("@Qid", newQid);
            try
            {
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex) { return false; }
        }
        public static void InsertTestInfo(SqlConnection gc, string newTestId, string grade, string subject, int newQid)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into TestInfo(TestId, Grade, Subject, Qid) values(@TestId, @Grade, @Subject, @Qid)";
            cmd.Parameters.AddWithValue("@TestId", newTestId);
            cmd.Parameters.AddWithValue("@Grade", grade);
            cmd.Parameters.AddWithValue("@Subject", subject);
            cmd.Parameters.AddWithValue("@Qid", newQid);
            cmd.ExecuteNonQuery();
        }
        public static void InsertIsHeadingUpdate(SqlConnection gc, string fileId, bool status)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into IsHeadingUpdate (FileId, IsUpdate) values (@FileId, @IsUpdate)";
            cmd.Parameters.AddWithValue("@FileId", fileId);
            cmd.Parameters.AddWithValue("@IsUpdate", status);
            cmd.ExecuteNonQuery();
        }
        public static void BackUpQuestionFile(SqlConnection gc, int grade, string subject, string fileName, byte[] fileBytes)
        {
            SqlConnection con = gc;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "insert into UploadedQuestionFile (FileId, Grade, Subject, FileTittle, WordBytes) values (@FileId, @Grade, @Subject, @FileTittle, @WordBytes)";
            cmd.Parameters.AddWithValue("@FileId", GetNewQuestionFileId());
            cmd.Parameters.AddWithValue("@Grade", grade);
            cmd.Parameters.AddWithValue("@Subject", subject);
            cmd.Parameters.AddWithValue("@FileTittle", fileName);
            cmd.Parameters.AddWithValue("@WordBytes", fileBytes);
            cmd.ExecuteNonQuery();
        }


        static string RemoveEndDot(string givenString)
        {
            try
            {
                if (givenString.Length > 0)
                {
                    if (givenString[givenString.Length - 1] == '.')
                        return givenString.Substring(0, givenString.Length - 2).Trim();
                    else
                        return givenString.Trim();
                }
                else
                    return givenString;
            }
            catch (Exception ex) 
            {
                return givenString;
            }
        }
    }
}
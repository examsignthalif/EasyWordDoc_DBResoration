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
        static void Main(string[] args)
        {
            List<TestModel> TestQuestionList = new List<TestModel>();
            List<QuestionModel> QList = WordProcessRepo.GetAllQuestion("00077");

            WordTools tool = new WordTools();

            // Generate as word file.
            List<FileItem> CurrentTestPath = new List<FileItem>();
            foreach(QuestionModel item in QList)
            {
                // after generated word record the file path to convert as xpsDocument.
                if (item.HasImage)
                    CurrentTestPath.Add(tool.Generate_WordDoc_Image(item));
                else
                    CurrentTestPath.Add(tool.Generate_WordDoc(item));
            }
        }
    }
}

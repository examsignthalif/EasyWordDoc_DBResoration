using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Model
{
    public class QuestionModel
    {
        public int Qid { get; set; }
        public string Question { get; set; }
        public bool HasImage { get; set; }

        public string Subject { get; set; }
        public string Grade { get; set; }
        public string QType { get; set; }

        public int Mark { get; set; }
        public string Heading { get; set; }
        public string Topic { get; set; }
        public string SheetType { get; set; }
        public List<ImageModel> ImageList { get; set; }
        public byte[] XpsByteData { get; set; }

        public QuestionModel()
        {
            ImageList = new List<ImageModel>();
        }
    }
}

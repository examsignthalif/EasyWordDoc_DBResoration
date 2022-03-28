using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Models
{
    public class OldQuesionModel
    {
        public string Question { get; set; }
        public bool HasImage { get; set; }

        public string Subject { get; set; }
        public string Grade { get; set; }
        public string QType { get; set; }

        public int Mark { get; set; }
        public string Heading { get; set; }
        public string Topic { get; set; }
        public string SheetType { get; set; }
        public List<OldImageModel> ImageList { get; set; }

        public OldQuesionModel()
        {
            ImageList = new List<OldImageModel>();
        }
    }
}

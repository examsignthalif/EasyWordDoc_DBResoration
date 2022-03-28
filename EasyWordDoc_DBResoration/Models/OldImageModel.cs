using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Models
{
    public class OldImageModel
    {
        public int Qid { get; set; }
        public string ImageNumber { get; set; }
        public byte[] ImageByteData { get; set; }
    }
}

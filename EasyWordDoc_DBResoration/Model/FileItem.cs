using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWordDoc_DBResoration.Model
{
    public class FileItem
    {
        public int Qno { get; set; }
        public string FilePath { get; set; }
        public int Marks { get; set; }
        public XpsDocument XPS { get; set; }
        public FileItem(int qno, string filePath)
        {
            this.Qno = qno;
            this.FilePath = filePath;
        }
        public FileItem()
        {

        }
    }
}

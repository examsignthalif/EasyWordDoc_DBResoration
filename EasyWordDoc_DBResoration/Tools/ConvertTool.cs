using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace EasyWordDoc_DBResoration.Tools
{
    public class ConvertTool
    {
        public static string Convert_WordToXPS(string word_Path)
        {
            string xps_Destination = word_Path.Replace(".docx", ".xps");
            ApplicationClass wordApp = new ApplicationClass();
            Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Open(word_Path);
                wordDoc.SaveAs2(xps_Destination, WdSaveFormat.wdFormatXPS);
                wordDoc.Close();
                wordApp.Quit();
            }
            catch (Exception ex) { }
            return xps_Destination;
        }
        public static byte[] ConvertWordToBytes(string filePath)
        {
            byte[] byteData = File.ReadAllBytes(filePath);
            return byteData;
        }
    }
}

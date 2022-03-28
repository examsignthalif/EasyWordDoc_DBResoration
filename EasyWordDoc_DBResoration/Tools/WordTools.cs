using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Documents;
using System.Xml;
using System.Xml.Linq;

using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Diagnostics;
using EasyWordDoc_DBResoration.Model;

namespace EasyWordDoc_DBResoration.Tools
{
    public class WordTools
    {
        static string ConnectionString = "Data Source=.;Initial Catalog=WordProcessNew;Integrated Security=True";
        public WordTools()
        {
            con = new SqlConnection(ConnectionString);
            cmd.Connection = con;
            con.Open();

            UnZipInto = RootFolder + w_UnzippedPath;
            MediaFile_Dest = UnZipInto + w_MediaPath;
            RelsFile_Dest = UnZipInto + w_DocRelsPath;
            XmlDocFile_Dest = UnZipInto + w_docXmlPath;

            ImageDB_Path = RootFolder + "Images\\";
            FINALOUTPUTFOLDER = RootFolder + "FINAL\\";
            CreateFolders();
        }

        SqlConnection con;
        SqlCommand cmd = new SqlCommand();
        // For proccess WORD file
        string Prefix_String = string.Empty;
        string InterStart_String = string.Empty;
        string InterEnd_String = string.Empty;
        string Suffix_String = string.Empty;
        string Shape_String = string.Empty;
        string NonShaped_String = string.Empty;
        int rid = 10;

        static string RootFolder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WordTesting\\OutputFolder\\");

        // Word file paths
        static string w_UnzippedPath = "Questions.docx.unzipped\\";

        static string w_MediaPath = "word\\media\\";
        static string w_DocRelsPath = "word\\_rels\\document.xml.rels";
        static string w_docXmlPath = "word\\document.xml";

        static string UnZipInto = string.Empty;
        static string ImageDB_Path = string.Empty;
        static string MediaFile_Dest = string.Empty;
        static string RelsFile_Dest = string.Empty;
        static string XmlDocFile_Dest = string.Empty;

        public static string FINALOUTPUTFOLDER = string.Empty;

        int sClass = 0;
        string sSubject = string.Empty;
        string sSheetType = string.Empty;
        string FileName = string.Empty;

        public FileItem Generate_WordDoc(QuestionModel questionItem)
        {
            GenerateDirectory(UnZipInto);

            Prefix_String = Get_EmptyPreFix();
            Suffix_String = Get_Suffix();
            InterStart_String = Get_InterStart();
            InterEnd_String = Get_InterEnd();
            NonShaped_String = Get_NonShape();
            Shape_String = Get_Shaped(questionItem.QType);

            StringBuilder sb = new StringBuilder();
            sb.Append(Prefix_String);
            sb.Append(InterStart_String);

            if (questionItem.ToString().ToUpper() == "Q" || questionItem.ToString().ToUpper() == "T")
                sb.Append(NonShaped_String);
            else
                sb.Append(Shape_String);

            sb.Append(questionItem.Question);
            sb.Append(InterEnd_String);
            sb.Append(Suffix_String);

            TOSAVE(XmlDocFile_Dest, sb.ToString());

            string fullPath = ResolveErrorandSaveDocument(questionItem.Qid.ToString());
            FindPageType(fullPath);
            return new FileItem(questionItem.Qid, fullPath);
        }
        public FileItem Generate_WordDoc_Image(QuestionModel questionItem)
        {
            GenerateDirectory(UnZipInto);

            Prefix_String = Get_EmptyPreFix();
            Suffix_String = Get_Suffix();
            InterStart_String = Get_InterStart();
            InterEnd_String = Get_InterEnd();
            NonShaped_String = Get_NonShape();
            Shape_String = Get_Shaped(questionItem.QType);


            StringBuilder sb = new StringBuilder();
            sb.Append(Prefix_String);
            sb.Append(InterStart_String);

            string Question_ItemString = questionItem.Question;
            for (int i = 0; i < questionItem.ImageList.Count(); i++)
            {
                rid++;
                string New_String_Replaced_Stuff = Question_ItemString.ToString().Replace("pic:cNvPr", "piccNvPr").Replace("a:blip", "ablip").Replace("r:embed", "rembed").Replace("w:tc", "wzzztc");

                XDocument xdocument = XDocument.Parse(New_String_Replaced_Stuff);
                var testelement = xdocument.Descendants("piccNvPr").Attributes("name").ToList();
                var element = xdocument.Descendants("piccNvPr").Attributes("name").ToList()[i];
                element.Value = questionItem.ImageList[i].ImageNumber + ".jpg";

                var testelement1 = xdocument.Descendants("ablip").Attributes("rembed").ToList();
                var element1 = xdocument.Descendants("ablip").Attributes("rembed").ToList()[i];
                element1.Value = "rId" + rid;

                var doc1 = xdocument.ToString().Replace("piccNvPr", "pic:cNvPr").Replace("ablip", "a:blip").Replace("rembed", "r:embed").Replace("wzzztc", "w:tc");
                Question_ItemString = string.Empty;
                Question_ItemString = doc1.ToString();

                ByteArrayToFile(MediaFile_Dest, questionItem.ImageList[i].ImageNumber, questionItem.ImageList[i].ImageByteData);
                AddImageResourceIdAndNameInDocXmlRels(RelsFile_Dest, questionItem.ImageList[i].ImageNumber, "rId" + rid);
            }

            if (questionItem.ToString().ToUpper() == "Q" || questionItem.ToString().ToUpper() == "T")
                sb.Append(NonShaped_String);
            else
                sb.Append(Shape_String);

            sb.Append(Question_ItemString);
            sb.Append(InterEnd_String);
            sb.Append(Suffix_String);

            TOSAVE(XmlDocFile_Dest, sb.ToString());


            string fullPath = ResolveErrorandSaveDocument(questionItem.Qid.ToString());
            FindPageType(fullPath);
            return new FileItem(questionItem.Qid, fullPath);
        }



        // Essentials
        public string GetCurrentTestId(int existcount)
        {
            int existtestcountlength = Convert.ToInt32(Math.Floor(Math.Log10(existcount) + 1));
            string testid = "00000";
            if (existcount == 0)
            {
                testid = "00001";
            }
            else if (existtestcountlength == 1)
            {
                testid = "0000" + existcount.ToString();
            }
            else if (existtestcountlength == 2)
            {
                testid = "000" + existcount.ToString();
            }
            else if (existtestcountlength == 3)
            {
                testid = "00" + existcount.ToString();
            }
            else if (existtestcountlength == 4)
            {
                testid = "0" + existcount.ToString();
            }
            else if (existtestcountlength == 5)
            {
                testid = existcount.ToString();
            }
            else
            {
                return testid;
            }
            return testid;
        }
        void GenerateDirectory(string destinationDirectorypath)
        {
            if (Directory.Exists(destinationDirectorypath))
            {
                Directory.Delete(destinationDirectorypath, true);
                Directory.CreateDirectory(destinationDirectorypath);
                CreateDefaultFoldersAndFileswithHeader(destinationDirectorypath);
            }
            else
            {
                Directory.CreateDirectory(destinationDirectorypath);
                CreateDefaultFoldersAndFileswithHeader(destinationDirectorypath);
            }
        }
        void CreateDefaultFoldersAndFileswithHeader(string Folderpath)
        {
            string relsFolder = Folderpath + "\\_rels";
            string docProsFolder = Folderpath + "\\docProps";
            string wordProsFolder = Folderpath + "\\word";
            string customxmlFolder = Folderpath + "\\customXml";
            cmd.CommandText = "select ContentTypeFile from DefaultFilesTable1";
            var contenttypefile = cmd.ExecuteScalar().ToString();
            XmlDocument docctype = new XmlDocument();
            docctype.LoadXml(contenttypefile);
            docctype.Save(Folderpath + "\\[Content_Types].xml");
            //create _rels folder
            if (!Directory.Exists(relsFolder))
            {
                Directory.CreateDirectory(relsFolder);
            }
            else
            {
                Directory.Delete(relsFolder, true);
                Directory.CreateDirectory(relsFolder);
            }
            //create docProsFolder
            if (!Directory.Exists(docProsFolder))
            {
                Directory.CreateDirectory(docProsFolder);
            }
            else
            {
                Directory.Delete(docProsFolder, true);
                Directory.CreateDirectory(docProsFolder);
            }
            //create wordProsFolder
            if (!Directory.Exists(wordProsFolder))
            {
                Directory.CreateDirectory(wordProsFolder);
            }
            else
            {
                Directory.Delete(wordProsFolder, true);
                Directory.CreateDirectory(wordProsFolder);
            }
            //create CustomxmlFolder
            if (!Directory.Exists(customxmlFolder))
            {
                Directory.CreateDirectory(customxmlFolder);
            }
            else
            {
                Directory.Delete(customxmlFolder, true);
                Directory.CreateDirectory(customxmlFolder);
            }
            // create _rels file
            cmd.CommandText = "select RelsFile from DefaultFilesTable1";
            var RelsFile = cmd.ExecuteScalar().ToString();
            XmlDocument docrelsfile = new XmlDocument();
            docrelsfile.LoadXml(RelsFile);
            docrelsfile.Save(relsFolder + "\\.rels");
            //create rels folder in customxml folder
            string custom_rels_folder = customxmlFolder + "\\" + "_rels";
            if (!Directory.Exists(custom_rels_folder))
            {
                Directory.CreateDirectory(custom_rels_folder);
            }
            else
            {
                Directory.Delete(custom_rels_folder, true);
                Directory.CreateDirectory(custom_rels_folder);
            }

            //create item1xms.rels file in custom xml rels folder
            cmd.CommandText = "select CustomRelsItem from DefaultFilesTable1";
            var customrelsitem = cmd.ExecuteScalar().ToString();
            XmlDocument customitem = new XmlDocument();
            customitem.LoadXml(customrelsitem);
            customitem.Save(custom_rels_folder + "\\item1.xml.rels");
            //create item1.xml file in customxmlfolder
            cmd.CommandText = "select ItemFile from DefaultFilesTable1";
            var ItemFile = cmd.ExecuteScalar().ToString();
            XmlDocument itemfiledoc = new XmlDocument();
            itemfiledoc.LoadXml(ItemFile);
            itemfiledoc.Save(customxmlFolder + "\\item1.xml");
            //create itemProps1.xml file in customxmlfolder
            cmd.CommandText = "select ItemProsFile from DefaultFilesTable1";
            var ItemProsFile = cmd.ExecuteScalar().ToString();
            XmlDocument ItemProsFiledoc = new XmlDocument();
            ItemProsFiledoc.LoadXml(ItemProsFile);
            ItemProsFiledoc.Save(customxmlFolder + "\\itemProps1.xml");
            //create files in docpros folder
            cmd.CommandText = "select AppFile from DefaultFilesTable1";
            var appFile = cmd.ExecuteScalar().ToString();
            XmlDocument docappfile = new XmlDocument();
            docappfile.LoadXml(appFile);
            docappfile.Save(docProsFolder + "\\app.xml");
            //create second file in docpros folder
            cmd.CommandText = "select CoreFile from DefaultFilesTable1";
            var coreFile = cmd.ExecuteScalar().ToString();
            XmlDocument doccorefile = new XmlDocument();
            docappfile.LoadXml(coreFile);
            docappfile.Save(docProsFolder + "\\core.xml");
            //create _rels folder in Word folder
            string word_relsFolder = wordProsFolder + "\\" + "_rels";
            if (!Directory.Exists(word_relsFolder))
            {
                Directory.CreateDirectory(word_relsFolder);
            }
            else
            {
                Directory.Delete(word_relsFolder, true);
                Directory.CreateDirectory(word_relsFolder);
            }
            //create theme folder in word folder
            string word_themeFolder = wordProsFolder + "\\" + "theme";
            if (!Directory.Exists(word_themeFolder))
            {
                Directory.CreateDirectory(word_themeFolder);
            }
            else
            {
                Directory.Delete(word_themeFolder, true);
                Directory.CreateDirectory(word_themeFolder);
            }
            //create media folder in word folder
            string word_mediaFolder = wordProsFolder + "\\" + "media";
            if (!Directory.Exists(word_mediaFolder))
            {
                Directory.CreateDirectory(word_mediaFolder);
            }
            else
            {
                Directory.Delete(word_mediaFolder, true);
                Directory.CreateDirectory(word_mediaFolder);
            }
            //create fonttablefile in word folder
            cmd.CommandText = "select FontTableFile from DefaultFilesTable1";
            var fonttableFile = cmd.ExecuteScalar().ToString();
            XmlDocument docfonttablefile = new XmlDocument();
            docfonttablefile.LoadXml(fonttableFile);
            docfonttablefile.Save(wordProsFolder + "\\fontTable.xml");
            //create settingsfile in word folder
            cmd.CommandText = "select settingsFile from DefaultFilesTable1";
            var settingsFile = cmd.ExecuteScalar().ToString();
            XmlDocument docsettings = new XmlDocument();
            docsettings.LoadXml(settingsFile);
            docsettings.Save(wordProsFolder + "\\settings.xml");
            //create stylefile in word folder
            cmd.CommandText = "select stylesfile from DefaultFilesTable1";
            var styleFile = cmd.ExecuteScalar().ToString();
            XmlDocument docstyle = new XmlDocument();
            docstyle.LoadXml(styleFile);
            docstyle.Save(wordProsFolder + "\\styles.xml");
            //create websettings in word folder
            cmd.CommandText = "select WebsettingsFile from DefaultFilesTable1";
            var websettingFile = cmd.ExecuteScalar().ToString();
            XmlDocument websettingstyle = new XmlDocument();
            websettingstyle.LoadXml(websettingFile);
            websettingstyle.Save(wordProsFolder + "\\webSettings.xml");
            //create endnotes.xml in word folder
            cmd.CommandText = "select Endnotes from DefaultFilesTable1";
            var Endnotes = cmd.ExecuteScalar().ToString();
            XmlDocument Endnotesdoc = new XmlDocument();
            Endnotesdoc.LoadXml(Endnotes);
            Endnotesdoc.Save(wordProsFolder + "\\endnotes.xml");
            //create Footnotes.xml in word folder
            cmd.CommandText = "select Footnotes from DefaultFilesTable1";
            var Footnotes = cmd.ExecuteScalar().ToString();
            XmlDocument Footnotesdoc = new XmlDocument();
            Footnotesdoc.LoadXml(Footnotes);
            Footnotesdoc.Save(wordProsFolder + "\\footnotes.xml");
            //create header1.xml in word folder

            //cmd.CommandText = "select Header1File from DefaultFilesTable1";
            //var HeaderFile = cmd.ExecuteScalar().ToString();            
            //var Header1File = ChangeClassSecionDetails(HeaderFile);\
            //XmlDocument Header1Filedoc = new XmlDocument();
            //Header1Filedoc.LoadXml(Header1File);
            //Header1Filedoc.Save(wordProsFolder + "\\header1.xml");



            //create numbering.xml in word folder
            cmd.CommandText = "select Header1File from DefaultFilesTable1";
            var Numbering1File = cmd.ExecuteScalar().ToString();
            XmlDocument Numbering1Filedoc = new XmlDocument();
            Numbering1Filedoc.LoadXml(Numbering1File);
            Numbering1Filedoc.Save(wordProsFolder + "\\numbering.xml");
            //create file in word _rels folder
            cmd.CommandText = "select DocRelsFile from DefaultFilesTable1";
            var docrelsFile = cmd.ExecuteScalar().ToString();
            XmlDocument xdocrelsfile = new XmlDocument();
            xdocrelsfile.LoadXml(docrelsFile);
            xdocrelsfile.Save(word_relsFolder + "\\document.xml.rels");
            //create header1.xml.rels file in _rels Folder
            cmd.CommandText = "select HeaderRels from DefaultFilesTable1";
            var HeaderRels = cmd.ExecuteScalar().ToString();
            XmlDocument HeaderRelsdoc = new XmlDocument();
            HeaderRelsdoc.LoadXml(HeaderRels);
            HeaderRelsdoc.Save(word_relsFolder + "\\header1.xml.rels");


            //create file in word theme folder
            cmd.CommandText = "select ThemeFile from DefaultFilesTable1";
            var themeFile = cmd.ExecuteScalar().ToString();
            XmlDocument xdocthemefile = new XmlDocument();
            xdocthemefile.LoadXml(themeFile);
            xdocthemefile.Save(word_themeFolder + "\\theme1.xml");
        }
        string ChangeClassSecionDetails(string headerxmlfile)
        {
            string sclass = "";
            string subject = "";
            string activity = "";
            if (sClass != 0)
            {
                sclass = sClass.ToString();
            }
            if (sSubject != "")
            {
                subject = sSubject;
            }
            if (sSheetType != "")
            {
                activity = sSheetType;
            }
            //return headerxmlfile.Replace("Class:", "Class:   " + sclass + "").Replace("Section:", " ").Replace("Subject:", "Subject: " + subject + "").Replace("Activity:", "" + activity + "   ").ToString();
            return headerxmlfile.Replace("Class:", string.Empty).Replace("Section:", string.Empty).Replace("Subject:", string.Empty).Replace("Activity:", string.Empty).ToString();
        }
        void AddImageResourceIdAndNameInDocXmlRels(string documentxmlrels, string imagename, string rid)
        {
            try
            {
                string newstring = documentxmlrels.ToString();
                XDocument xdocument = XDocument.Load(newstring);
                XElement newTag1 = new XElement("Relationship",
                new XAttribute("Id", rid),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                new XAttribute("Target", "media/" + imagename + ".jpg"));
                xdocument.Root.LastNode.AddAfterSelf(newTag1);
                xdocument.Save(documentxmlrels);
            }
            catch (Exception ex)
            {
            }
        }
        void TOSAVE(string docFilePath, string xmlData)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlData.ToString());
            doc.Save(docFilePath);
        }
        static string ResolveErrorandSaveDocument(string filename = "Question")
        {
            //string outputfilepath = filename + ques_id + " .docx";
            if (!Directory.Exists(FINALOUTPUTFOLDER))
            {
                Directory.CreateDirectory(FINALOUTPUTFOLDER);
            }
            if (!filename.Contains(".docx")) filename += ".docx";
            if (!filename.Contains("Question")) filename = "Question" + filename;

            string outputfilepath = FINALOUTPUTFOLDER + filename;
            ConvertToZipFile(UnZipInto, outputfilepath);

            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(outputfilepath, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, true, ref oMissing, ref oMissing, ref oMissing);
            doc.SaveAs2(outputfilepath, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            doc.Close();
            app.Quit();

            if (Directory.Exists(UnZipInto))
                Directory.Delete(UnZipInto, true);
            return outputfilepath;
        }
        static void FindPageType(string filename)
        {
            //string filename = @"D:\WordTesting\OutputFolder\Question13.docx";
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filename, true))
            {
                string docText = null;
                int pageCount = Convert.ToInt32(wordDoc.ExtendedFilePropertiesPart.Properties.Pages.Text);
                if (pageCount == 1)
                {
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd().Replace("#S~M~C", "");
                    }
                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
                else
                {
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd().Replace("#S~M~C", "##");
                    }
                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
            }
        }
        void CreateFolders()
        {
            if (!Directory.Exists(RootFolder))
            {
                Directory.CreateDirectory(RootFolder);
            }
            if (!Directory.Exists(ImageDB_Path))
            {
                try
                {
                    var currentfile = AppDomain.CurrentDomain.BaseDirectory + "\\Img\\image1.png";
                    var desfile = ImageDB_Path + "\\image1.png";
                    Directory.CreateDirectory(ImageDB_Path);
                    File.Copy(currentfile, desfile);
                }
                catch (Exception ex) { }
            }
        }

        // Conversions
        void ConvertToUnZip(string filename)
        {
            using (ZipArchive archive = ZipFile.OpenRead(filename))
            {
                string contextPath = filename + ".unzipped";
                if (Directory.Exists(contextPath))
                    Directory.Delete(contextPath, true);
                archive.ExtractToDirectory(filename + ".unzipped");
            }
        }
        void Deleteunzippedfolder(string filename)
        {
            string filefullpath = filename + ".unzipped";
            if (Directory.Exists(filefullpath))
            {
                Directory.Delete(filefullpath, true);
            }
        }
        public static byte[] converterDemo(string x)
        {
            FileStream fs = new FileStream(x, FileMode.Open, FileAccess.Read);
            byte[] imgByteArr = new byte[fs.Length];
            fs.Read(imgByteArr, 0, Convert.ToInt32(fs.Length));
            fs.Close();
            return imgByteArr;
        }
        private void ByteArrayToFile(string bPath, string fName, byte[] content)
        {
            //Save the Byte Array as File.
            string filePath = bPath + fName + ".jpg";
            File.WriteAllBytes(filePath, content);
        }
        static void ConvertToZipFile(string filename, string destinationFileName)
        {
            try
            {
                if (File.Exists(destinationFileName))
                {
                    File.Delete(destinationFileName);
                }
                ZipFile.CreateFromDirectory(filename, destinationFileName);
            }
            catch (Exception ex)
            { }
        }
        string Addtablecell(string inputstring)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<w:tc>");
            sb.Append(inputstring);
            sb.Append("</w:tc>");
            return sb.ToString();
        }

        // Needs
        static string Get_EmptyPreFix()
        {
            var prefixtext = File.ReadAllText(MyDirectory() + @"\prefix.txt");
            return prefixtext.ToString();
        }
        static string Get_TestHeaderPreFix()
        {
            var prefixtext = File.ReadAllText(MyDirectory() + @"\prefix12NewFormat.txt");
            return prefixtext.ToString();
        }
        static string Get_Suffix()
        {
            var suf = File.ReadAllLines(MyDirectory() + @"\suffix1.txt").ToList();
            StringBuilder ssb = new StringBuilder();
            foreach (var n in suf)
            {
                ssb.Append(n.ToString());
            }
            return ssb.ToString();
        }
        static string Get_InterStart()
        {
            string interstart = string.Empty;
            var s1 = File.ReadAllLines(MyDirectory() + @"\inter1.txt").ToList();
            foreach (var n in s1)
            {
                interstart = n.ToString();
            }
            return interstart;
        }
        static string Get_InterEnd()
        {
            string interend = string.Empty;
            var s2 = File.ReadAllLines(MyDirectory() + @"\inter2.txt").ToList();
            foreach (var n in s2)
            {
                interend = n.ToString();
            }
            return interend;
        }
        static string Get_Shaped(string Qtype)
        {
            List<string> s3 = new List<string>();
            if (Qtype.ToUpper() == "single".ToUpper())
            {
                s3 = File.ReadAllLines(MyDirectory() + @"\shape2Rectangle.txt").ToList();
            }
            else if (Qtype.ToUpper() == "multiple".ToUpper())
            {
                s3 = File.ReadAllLines(MyDirectory() + @"\shape6Rectangle.txt").ToList();
            }

            //Set Shaped file
            string shapefile = string.Empty;
            foreach (var n in s3)
            {
                shapefile = n.ToString();
            }
            return shapefile;
        }
        static string Get_NonShape()
        {
            var s4 = File.ReadAllLines(MyDirectory() + @"\nonshape.txt").ToList();
            string nonshape = string.Empty;
            foreach (var n in s4)
            {
                nonshape = n.ToString();
            }
            return nonshape;
        }
        static string MyDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
    }
}

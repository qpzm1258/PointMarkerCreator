using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace wordtest
{
    class Program
    {
        private static Dictionary<string, Range> m_RangeMap;
        private static Dictionary<string, InlineShape> m_InlineShapeMap;
        private static void SetText(ref _Document doc, string bookmark, string value, string fontName = "Times New Roman",float fontSize=10.5F)
        {
            if (doc.Bookmarks.Exists(bookmark))
            {
                m_RangeMap[bookmark].Font.Name = fontName;
                m_RangeMap[bookmark].Font.Size = fontSize;
                m_RangeMap[bookmark].Text = value;
            }
        }

        public static void InsertPicture(ref _Document doc, string bookmark, string picturePath, float width, float hight)
        {
            object miss = System.Reflection.Missing.Value;
            object oStart = bookmark;
            Object linkToFile = false;       //图片是否为外部链接
            Object saveWithDocument = true;  //图片是否随文档一起保存 
            object range = m_RangeMap[bookmark];//图片插入位置
            if (m_InlineShapeMap[bookmark] != null)
            {
                m_InlineShapeMap[bookmark].Delete();
                m_InlineShapeMap[bookmark] = null;
            }
            m_InlineShapeMap[bookmark]=doc.InlineShapes.AddPicture(picturePath, ref linkToFile, ref saveWithDocument, ref range);
            m_InlineShapeMap[bookmark].Width = width;   //设置图片宽度
            m_InlineShapeMap[bookmark].Height = hight;  //设置图片高度
        }



        private static void CreateDocument()
        {
            try
            {
                string path = System.Environment.CurrentDirectory;
                //Create an instance for word app
                Microsoft.Office.Interop.Word._Application winword = new Microsoft.Office.Interop.Word.Application();
                //Set animation status for word application
                //winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.
                winword.Visible = false;

                //Create a missing variable for missing value
                object missing = System.Reflection.Missing.Value;
                object tempflie = path+"\\point\\template.dotx";
                //Create a new document from document
                Microsoft.Office.Interop.Word._Document document = winword.Documents.Add(tempflie, ref missing, ref missing, ref missing);
                document.SpellingChecked = false;
                document.ShowSpellingErrors = false;
                m_RangeMap = new Dictionary<string, Range>();
                m_InlineShapeMap = new Dictionary<string, InlineShape>();
                Console.WriteLine("Environment Setting:");

                string input;
                do
                {
                    Console.WriteLine("[Debug Model]");
                    Console.Write("Do you want the program work on debug model? y(es) or n(o)");
                    input = Console.ReadLine().ToLower();  
                } while (input != "y" && input != "n");

                //winword.Visible = (input == "y");

                Console.WriteLine("[Txt Setting]");
                Console.WriteLine("Please put txt into path: " + path + "\\point\\point.txt");
                Console.WriteLine("Txt Format example: [PointName],[X],[Y],[H]");
                Console.WriteLine();
                Console.WriteLine("[Image Setting]");
                Console.WriteLine("Please put big image into path: " + path + "\\point\\imagedir\\big\\");
                Console.WriteLine("Please put middle image into path: " + path+"\\point\\imagedir\\middle\\");
                Console.WriteLine("Please put small image into path: " + path + "\\point\\imagedir\\small\\");
                Console.WriteLine("Image in each directory must be named: [PointName].jpg");
                Console.WriteLine();
                Console.WriteLine("[Document Setting]");
                Console.WriteLine("Document will be created into path: " + path + "\\point\\point.docx");
                Console.WriteLine("Warning !!! Check if the document exist, it will be covered");
                Console.WriteLine();
                Console.WriteLine("Please check your files in the right path and press any key to continue...");
                Console.ReadKey();

                //Init bookmarks map
                foreach (Bookmark bookmark in document.Bookmarks)
                {
                    m_RangeMap.Add(bookmark.Name, bookmark.Range);
                    m_InlineShapeMap.Add(bookmark.Name, null);
                }
                
                //Read txt
                StreamReader sr = new StreamReader(path+"\\point\\point.txt", Encoding.Default);
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    Console.WriteLine(line.ToString());
                    string[] values=line.Split(new char[1]{','});
                    if (values.Count() == 4)
                    {
                        //Set value to bookmark
                        SetText(ref document, "PointName", values[0]);
                        SetText(ref document, "X", values[1]);
                        SetText(ref document, "Y", values[2]);
                        SetText(ref document, "H", values[3]);
                        InsertPicture(ref document, "BigImage", path + "\\point\\imagedir\\big\\" + values[0] + ".jpg", 239F, 360F);
                        InsertPicture(ref document, "MiddleImage", path + "\\point\\imagedir\\middle\\" + values[0] + ".jpg", 194F, 274F);
                        InsertPicture(ref document, "SmallImage", path + "\\point\\imagedir\\small\\" + values[0] + ".jpg", 176F,193F);
                        Table firstTable = document.Tables[1];
                        firstTable.Range.Copy();
                        //goto next page
                        document.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                        document.Content.Paragraphs.Add(ref missing);
                        Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                        para1.Range.Paste();
                    }
                }

                //delete template table
                int num = document.ComputeStatistics(WdStatistic.wdStatisticPages,ref missing);
                var app = new Application();
                document.Application.Selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, num, 2);
                document.Application.Selection.Bookmarks[@"\Page"].Select();
                document.Application.Selection.Delete();  

                //Save the document
                object filename = path + "\\point\\point.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                Console.WriteLine("Document created successfully !");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            killWinWordProcess();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        /*
         * Kill WinWord Process
         */
        public static void killWinWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("This program base on .NET Framework 4.5 and Microsoft office 2010/2013/2016");
            killWinWordProcess();
            CreateDocument();
        }
    }
}

using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Processing document");

            //Find File
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filepath = Path.Combine(desktopPath + @"\WordFile\test_xml.docx");

            //Save Original File
            string originalFilename = Path.GetFileNameWithoutExtension(filepath);
            var originalFile = Path.Combine(desktopPath, @"WordFile\" + originalFilename + "-v1.docx");
            System.IO.File.Copy(filepath, originalFile);

            //Load File
            WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true);
            Body body = doc.MainDocumentPart.Document.Body;

            //Find Title Paragraph
            IEnumerable<Paragraph> paragraphs = body.Elements<Paragraph>();
            string title = "A method to work with Mri devices";
            Paragraph titleParagraph = null;
            foreach (Paragraph p in paragraphs)
            {
                string text = p.InnerText;
                if (text.Equals(title)) {
                    titleParagraph = p;
                }
            }

            //Replace Title Text
            string correctedTitle = "A Method to Work with MRI Devices";
            foreach (Run r in titleParagraph.Elements<Run>())
            {
                foreach(Text t in r.Elements<Text>())
                {
                    t.Text = "";
                }
            }
            Run run = titleParagraph.AppendChild(new Run());
            run.AppendChild(new Text(correctedTitle));


            //Close and Save File
            doc.Save();
            doc.Close();

            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}

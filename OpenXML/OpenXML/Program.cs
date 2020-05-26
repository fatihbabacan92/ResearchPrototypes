using System;

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
        }
    }
}

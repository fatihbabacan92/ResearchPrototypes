using System;
using System.Drawing;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace XceedWords
{
    class Program
    {
        static void Main(string[] args)
        {
            //find and replace text
            //Apply format rules

            Console.WriteLine("Processing document");

            //Find File from Desktop/WordFile
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filename = Path.Combine(desktopPath + @"\WordFile\test.docx");

            //Save Original File
            string originalFilename = Path.GetFileNameWithoutExtension(filename);
            var originalFile = Path.Combine(desktopPath, @"WordFile\" + originalFilename + "-v1.docx");
            System.IO.File.Copy(filename, originalFile);
            //Load Document -- DocX.Create() for creating new document
            var document = DocX.Load(filename);

            //Set Default Font to Times New Roman, 12px, Black
            document.SetDefaultFont(new Font("Times New Roman"), 12d, Color.Black);


        }
    }
}

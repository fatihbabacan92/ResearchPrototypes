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

            //Find File
            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            filename = Path.Combine(filename + @"\WordFile\test.docx");

            //Load Document -- DocX.Create() for creating new document
            var document = DocX.Load(filename);

            //Set Default Font to Times New Roman, 12px, Black
            document.SetDefaultFont(new Font("Times New Roman"), 12d, Color.Black);


        }
    }
}

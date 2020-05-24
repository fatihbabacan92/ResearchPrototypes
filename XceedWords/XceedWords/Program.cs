using System;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace XceedWords
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //Find File
            string filename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            filename = Path.Combine(filename + @"\WordFile\test.docx");
            //Load Document -- DocX.Create() for creating new document
            var document = DocX.Load(filename);
        }
    }
}

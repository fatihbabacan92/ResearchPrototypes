using System;
using System.IO;

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
        }
    }
}

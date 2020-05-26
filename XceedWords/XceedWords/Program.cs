using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
            using var document = DocX.Load(filename);

            //Set Default Font to Times New Roman, 12px, Black -- Useful when inserting new paragraphs
            document.SetDefaultFont(new Font("Times New Roman"), 12d, Color.Black);

            //Replace Title Text
            string title = "This is the Ai title";
            string correctedTitle = "This is the AI Title";
            document.ReplaceText(title, correctedTitle);

            //Find Title Paragraph
            var titleParagraph = document.Paragraphs.Where(p => p.Text.Equals(correctedTitle)).First();

            //Apply Format Rules
            titleParagraph.Bold(true);
            titleParagraph.FontSize(64d);
            titleParagraph.Color(Color.Black);

            //Misc.
            document.PageLayout.Orientation = Orientation.Portrait;

            //Save Documents
            document.SaveAs(filename);
            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}

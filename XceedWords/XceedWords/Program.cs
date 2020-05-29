using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
            Font times = new Font("Times New Roman");
            titleParagraph.Font(times);
            titleParagraph.FontSize(15);
            titleParagraph.Bold(true);
            titleParagraph.Bold(true);
            titleParagraph.UnderlineStyle(UnderlineStyle.none);
            titleParagraph.Italic(false);
            titleParagraph.Color(Color.Black);

            titleParagraph.Alignment = Alignment.center;
            titleParagraph.IndentationBefore = 0;
            titleParagraph.IndentationAfter = 0;
            titleParagraph.SpacingAfter(12);
            titleParagraph.SetLineSpacing(LineSpacingType.Line, 17);
            titleParagraph.KeepWithNextParagraph(false);
            titleParagraph.KeepLinesTogether(false);
            titleParagraph.IndentationFirstLine = 0;

            //Add table
            var t = document.AddTable(2, 2);
            t.Design = TableDesign.ColorfulListAccent1;
            t.Alignment = Alignment.center;
            t.Rows[0].Cells[0].Paragraphs[0].Append("Fatih");
            t.Rows[0].Cells[1].Paragraphs[0].Append("18/20");
            t.Rows[1].Cells[0].Paragraphs[0].Append("Kevin");
            t.Rows[1].Cells[1].Paragraphs[0].Append("10/20");

            //Misc.
            document.PageLayout.Orientation = Orientation.Portrait;

            //Save Documents
            document.SaveAs(filename);
            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}

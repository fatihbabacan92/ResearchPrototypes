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

            //Replace Keywords Text
            string keywords = "keywords: ai, titles, foNts; autoMation…";
            string correctedKeywords = "Keywords:\t AI, Titles, Fonts, Automation.";
            document.ReplaceText(keywords, correctedKeywords);

            //Repalce Abstract Text
            string abstractText = "Abstract: Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vestibulum egestas mi ante, vitae eleifend sapien pharetra sed. Praesent a sapien nibh. Fusce sagittis aliquet dignissim. Donec elementum odio dolor, sit amet mattis dui auctor suscipit. Fusce consectetur consectetur enim at ultricies. Suspendisse potenti. Morbi eget magna ac arcu placerat interdum sed sit amet neque. Cras sodales urna nibh, pretium pharetra ipsum sollicitudin non. Etiam interdum venenatis fringilla. Phasellus non augue sed magna laoreet mollis a id tellus. Ut tincidunt finibus neque, sed volutpat nunc rutrum tincidunt. Pellentesque orci massa, laoreet non velit ut, facilisis interdum lectus.";
            string correctedAbstractText = "Abstract:\t Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vestibulum egestas mi ante, vitae eleifend sapien pharetra sed. Praesent a sapien nibh. Fusce sagittis aliquet dignissim. Donec elementum odio dolor, sit amet mattis dui auctor suscipit. Fusce consectetur consectetur enim at ultricies. Suspendisse potenti. Morbi eget magna ac arcu placerat interdum sed sit amet neque. Cras sodales urna nibh, pretium pharetra ipsum sollicitudin non. Etiam interdum venenatis fringilla. Phasellus non augue sed magna laoreet mollis a id tellus. Ut tincidunt finibus neque, sed volutpat nunc rutrum tincidunt. Pellentesque orci massa, laoreet non velit ut, facilisis interdum lectus.";
            document.ReplaceText(abstractText, correctedAbstractText);

            //Find Paragraphs
            var titleParagraph = document.Paragraphs.Where(p => p.Text.Equals(correctedTitle)).First();
            var keywordsParagraph = document.Paragraphs.Where(p => p.Text.Equals(correctedKeywords)).First();
            var abstractParagraph = document.Paragraphs.Where(p => p.Text.Equals(correctedAbstractText)).First();

            //Apply Format Rules
            Font times = new Font("Times New Roman");
            titleParagraph.Font(times);
            titleParagraph.FontSize(15);
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

            keywordsParagraph.Font(times);
            keywordsParagraph.FontSize(11);
            keywordsParagraph.Color(Color.Black);
            keywordsParagraph.SpacingAfter(7);

            abstractParagraph.Font(times);
            abstractParagraph.FontSize(11);
            abstractParagraph.Color(Color.Black);
            abstractParagraph.SpacingAfter(7);
            abstractParagraph.Bold(false);

            //Add table
            var t = document.AddTable(2, 2);
            t.Design = TableDesign.ColorfulListAccent1;
            t.Alignment = Alignment.center;
            t.Rows[0].Cells[0].Paragraphs[0].Append("Name");
            t.Rows[0].Cells[1].Paragraphs[0].Append("Results");
            t.Rows[1].Cells[0].Paragraphs[0].Append("Kevin");
            t.Rows[1].Cells[1].Paragraphs[0].Append("9/20");
            abstractParagraph.InsertTableAfterSelf(t);


            //Misc.
            document.PageLayout.Orientation = Orientation.Portrait;

            //Save Documents
            document.SaveAs(filename);
            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}
